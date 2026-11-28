// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus
{
    public class FieldRetriever
    {
        public static string InstrText(XElement root, int id)
        {

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

#if false
            // This is the old code.  Both versions work - the caching version is significantly faster.
            var relevantElements = root.Descendants()
                .Where(e =>
                {
                    Stack<FieldElementTypeInfo> s = e.Annotation<Stack<FieldElementTypeInfo>>();
                    if (s != null)
                        return s.Any(z => z.Id == id &&
                            z.FieldElementType == FieldElementTypeEnum.InstrText);
                    return false;
                })
                .ToList();
#else
            var cachedAnnotationInformation = root.Annotation<Dictionary<int, List<XElement>>>();
            if (cachedAnnotationInformation == null)
                throw new DocxodusException("Internal error");

            // it is possible that a field code contains no instr text
            if (!cachedAnnotationInformation.ContainsKey(id))
                return "";
            var relevantElements = cachedAnnotationInformation[id];
#endif

            // Every element reachable here came from cachedAnnotationInformation, which
            // AnnotateWithFieldInfo populates only from elements that already carry a non-null
            // Stack<FieldElementTypeInfo> annotation — so the Annotation<>() calls below are
            // guaranteed non-null.
            var groupedSubFields = relevantElements
                .GroupAdjacent(e =>
                {
                    Stack<FieldElementTypeInfo> s = e.Annotation<Stack<FieldElementTypeInfo>>()!;
                    var stackElement = s.FirstOrDefault(z => z.Id == id);
                    var elementsBefore = s.TakeWhile(z => z != stackElement);
                    return elementsBefore.Any();
                })
                .ToList();

            // A deleted field (its begin character is in a w:del or w:moveFrom) keeps its code in
            // w:delInstrText, or in w:instrText inside a move source. A live field means only its
            // live code: one whose code was edited under tracking holds the old code as
            // w:delInstrText beside the new code as w:instrText, and the field means the new code.
            var deletedField = IsDeletedField(root, id);

            var instrText = groupedSubFields
                .Select(g =>
                {
                    if (g.Key == false)
                    {
                        return g.Select(e =>
                        {
                            Stack<FieldElementTypeInfo> s = e.Annotation<Stack<FieldElementTypeInfo>>()!;
                            // e is a member of its own stack's grouping, so a stack entry with
                            // Id == id (namely e's own entry) is guaranteed present.
                            var stackElement = s.FirstOrDefault(z => z.Id == id)!;
                            if (stackElement.FieldElementType == FieldElementTypeEnum.InstrText &&
                                (e.Name == w + "instrText" || (deletedField && e.Name == w + "delInstrText")))
                                return e.Value;
                            return "";
                        })
                            .StringConcatenate();
                    }
                    else
                    {
                        Stack<FieldElementTypeInfo> s = g.First().Annotation<Stack<FieldElementTypeInfo>>()!;
                        var stackElement = s.FirstOrDefault(z => z.Id == id);
                        var elementBefore = s.TakeWhile(z => z != stackElement).Last();
                        var subFieldId = elementBefore.Id;
                        return InstrText(root, subFieldId);
                    }
                })
                .StringConcatenate();

            return "{" + instrText + "}";
        }

        /// <summary>
        /// The code of field <paramref name="id"/> as Word evaluates it to run the field: the code's own text, with
        /// each field nested in it replaced by that field's cached result, so a <c>HYPERLINK</c> whose address is a
        /// <c>MERGEFIELD</c> points to the merged address. <see cref="InstrText"/> keeps a nested field's code instead,
        /// in braces. Only results show through: a field nested in a nested field's code adds nothing.
        /// </summary>
        internal static string EvaluatedInstrText(XElement root, int id)
        {
            var cached = root.Annotation<Dictionary<int, List<XElement>>>();
            if (cached == null || !cached.TryGetValue(id, out var elements))
                return "";
            // As InstrText reads a deleted field's code from w:delInstrText, a deleted field's nested results are
            // w:delText.
            var deletedField = IsDeletedField(root, id);
            var code = new StringBuilder();
            foreach (var element in elements)
            {
                bool nested = false, shown = true;
                foreach (var info in element.Annotation<Stack<FieldElementTypeInfo>>()!)
                {
                    if (info.Id == id)
                        break;
                    nested = true;
                    shown &= info.FieldElementType == FieldElementTypeEnum.Result;
                }
                if (nested ? shown && (element.Name == W.t || (deletedField && element.Name == W.delText))
                    : element.Name == W.instrText || (deletedField && element.Name == W.delInstrText))
                    code.Append(element.Value);
            }
            return code.ToString();
        }

        /// <summary>
        /// Returns whether <paramref name="element"/> is in the cached result of a field of
        /// <paramref name="fieldType"/>: a complex field's, or a simple field's, whose content is
        /// its result. Nested fields are considered independently, so a PAGEREF result nested
        /// inside a TOC result still reports both contexts correctly.
        /// </summary>
        internal static bool IsFieldResult(XElement element, string fieldType)
        {
            if (element == null || string.IsNullOrWhiteSpace(fieldType))
                return false;

            // Asked for every run, so it walks parents and the stack without allocating, and reads each field's type
            // once.
            var root = element;
            for (var ancestor = element.Parent; ancestor != null; ancestor = ancestor.Parent)
            {
                if (ancestor.Name == W.fldSimple
                    && string.Equals(FieldType((string?)ancestor.Attribute(W.instr) ?? ""), fieldType, StringComparison.OrdinalIgnoreCase))
                    return true;
                root = ancestor;
            }

            var stack = element.Annotation<Stack<FieldElementTypeInfo>>();
            if (stack == null)
                return false;
            FieldTypes? types = null;
            foreach (var info in stack)
            {
                if (info.FieldElementType != FieldElementTypeEnum.Result)
                    continue;
                if (types == null && (types = root.Annotation<FieldTypes>()) == null)
                    root.AddAnnotation(types = new FieldTypes());
                if (!types.TryGetValue(info.Id, out var type))
                    types[info.Id] = type = FieldType(root, info.Id);
                if (string.Equals(type, fieldType, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        /// <summary>A field instruction's type, its first word as written (<c>HYPERLINK</c>, <c>TOC</c>); Word reads it
        /// in any case.</summary>
        internal static string FieldType(string instruction)
        {
            int start = 0;
            while (start < instruction.Length && char.IsWhiteSpace(instruction[start])) start++;
            int end = start;
            while (end < instruction.Length && !char.IsWhiteSpace(instruction[end])) end++;
            return instruction.Substring(start, end - start);
        }

        /// <summary>The type of the field <paramref name="id"/>, as <see cref="FieldType(string)"/> reads it from
        /// <see cref="InstrText"/>, but from the code's first characters only: an instruction can be long, and
        /// <see cref="IsFieldResult"/> asks this for every field around every run.</summary>
        internal static string FieldType(XElement root, int id)
        {
            var cached = root.Annotation<Dictionary<int, List<XElement>>>();
            if (cached == null || !cached.TryGetValue(id, out var elements))
                return "";
            var deletedField = IsDeletedField(root, id);
            var word = new StringBuilder();
            foreach (var element in elements)
            {
                // The instruction's own code, as InstrText reads it; a nested field ends the first word.
                var own = element.Annotation<Stack<FieldElementTypeInfo>>()!.Peek();
                if (own.Id != id)
                    return word.ToString();
                if (own.FieldElementType != FieldElementTypeEnum.InstrText
                    || (element.Name != W.instrText && !(deletedField && element.Name == W.delInstrText)))
                    continue;
                foreach (var c in element.Value)
                {
                    if (!char.IsWhiteSpace(c))
                        word.Append(c);
                    else if (word.Length > 0)
                        return word.ToString();
                }
            }
            return word.ToString();
        }

        /// <summary>The type of each field of a part, by id, as <see cref="IsFieldResult"/> has read it.</summary>
        private sealed class FieldTypes : Dictionary<int, string> { }

        public static void AnnotateWithFieldInfo(OpenXmlPart part)
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            XElement root = part.GetXDocument().Root!;
            // A field result can be normalized between reads. Refresh the complete annotation
            // set so removed nested fields cannot leave stale ids or instruction-cache entries.
            root.RemoveAnnotations<Dictionary<int, List<XElement>>>();
            root.RemoveAnnotations<DeletedFieldIds>();
            root.RemoveAnnotations<FieldTypes>();
            foreach (var element in root.DescendantsAndSelf())
                element.RemoveAnnotations<Stack<FieldElementTypeInfo>>();
            // Each note and text box is a story of its own (a comment too, should its part be annotated), so a field left open in one (malformed
            // markup) cannot run into the next: the walk starts each with no field open, and resumes the
            // enclosing story's fields after a text box. Ids keep increasing across the part, since the
            // instruction cache is keyed by id on the part's root (issue #1876).
            var boxes = new Stack<(XElement Box, Stack<FieldElementTypeInfo>? Saved)>();
            var r = root.DescendantsAndSelf()
                .Rollup(
                    new FieldElementTypeStack
                    {
                        Id = 0,
                        FiStack = null,
                    },
                    (e, s) =>
                    {
                        while (boxes.Count > 0 && !e.Ancestors().Contains(boxes.Peek().Box))
                            s = new FieldElementTypeStack { Id = s.Id, FiStack = boxes.Pop().Saved };
                        if (e.Name == W.footnote || e.Name == W.endnote || e.Name == W.comment)
                            s = new FieldElementTypeStack { Id = s.Id, FiStack = null };
                        else if (e.Name == W.txbxContent)
                        {
                            boxes.Push((e, s.FiStack));
                            s = new FieldElementTypeStack { Id = s.Id, FiStack = null };
                        }

                        // w:fldChar always carries w:fldCharType per the WordprocessingML schema.
                        if (e.Name == w + "fldChar")
                        {
                            if (e.Attribute(w + "fldCharType")!.Value == "begin")
                            {
                                Stack<FieldElementTypeInfo> fis;
                                if (s.FiStack == null)
                                    fis = new Stack<FieldElementTypeInfo>();
                                else
                                    fis = new Stack<FieldElementTypeInfo>(s.FiStack.Reverse());
                                fis.Push(
                                    new FieldElementTypeInfo
                                    {
                                        Id = s.Id + 1,
                                        FieldElementType = FieldElementTypeEnum.Begin,
                                    });
                                return new FieldElementTypeStack
                                {
                                    Id = s.Id + 1,
                                    FiStack = fis,
                                };
                            };
                            // A separate or end with no field open (malformed markup, such as a stray end
                            // character in a header) belongs to no field: it leaves the state as it is, rather
                            // than failing the whole conversion (issue #1876).
                            if (s.FiStack is not { Count: > 0 })
                                return s;
                            if (e.Attribute(w + "fldCharType")!.Value == "separate")
                            {
                                Stack<FieldElementTypeInfo> fis = new Stack<FieldElementTypeInfo>(s.FiStack!.Reverse());
                                FieldElementTypeInfo wfi = fis.Pop();
                                fis.Push(
                                    new FieldElementTypeInfo
                                    {
                                        Id = wfi.Id,
                                        FieldElementType = FieldElementTypeEnum.Separate,
                                    });
                                return new FieldElementTypeStack
                                {
                                    Id = s.Id,
                                    FiStack = fis,
                                };
                            }
                            if (e.Attribute(w + "fldCharType")!.Value == "end")
                            {
                                Stack<FieldElementTypeInfo> fis = new Stack<FieldElementTypeInfo>(s.FiStack!.Reverse());
                                FieldElementTypeInfo wfi = fis.Pop();
                                return new FieldElementTypeStack
                                {
                                    Id = s.Id,
                                    FiStack = fis,
                                };
                            }
                        }
                        if (s.FiStack == null || s.FiStack.Count() == 0)
                            return s;
                        FieldElementTypeInfo wfi3 = s.FiStack.Peek();
                        if (wfi3.FieldElementType == FieldElementTypeEnum.Begin)
                        {
                            Stack<FieldElementTypeInfo> fis = new Stack<FieldElementTypeInfo>(s.FiStack.Reverse());
                            FieldElementTypeInfo wfi2 = fis.Pop();
                            fis.Push(
                                new FieldElementTypeInfo
                                {
                                    Id = wfi2.Id,
                                    FieldElementType = FieldElementTypeEnum.InstrText,
                                });
                            return new FieldElementTypeStack
                            {
                                Id = s.Id,
                                FiStack = fis,
                            };
                        }
                        if (wfi3.FieldElementType == FieldElementTypeEnum.Separate)
                        {
                            Stack<FieldElementTypeInfo> fis = new Stack<FieldElementTypeInfo>(s.FiStack.Reverse());
                            FieldElementTypeInfo wfi2 = fis.Pop();
                            fis.Push(
                                new FieldElementTypeInfo
                                {
                                    Id = wfi2.Id,
                                    FieldElementType = FieldElementTypeEnum.Result,
                                });
                            return new FieldElementTypeStack
                            {
                                Id = s.Id,
                                FiStack = fis,
                            };
                        }
                        if (wfi3.FieldElementType == FieldElementTypeEnum.End)
                        {
                            Stack<FieldElementTypeInfo>? fis = new Stack<FieldElementTypeInfo>(s.FiStack!.Reverse());
                            fis.Pop();
                            if (!fis.Any())
                                fis = null;
                            return new FieldElementTypeStack
                            {
                                Id = s.Id,
                                FiStack = fis,
                            };
                        }
                        return s;
                    });
            var elementPlusInfo = root.DescendantsAndSelf().PtZip(r, (t1, t2) =>
            {
                return new
                {
                    Element = t1,
                    Id = t2.Id,
                    WmlFieldInfoStack = t2.FiStack,
                };
            });
            foreach (var item in elementPlusInfo)
            {
                if (item.WmlFieldInfoStack != null)
                    item.Element.AddAnnotation(item.WmlFieldInfoStack);
            }

            //This code is useful when you want to take a look at the annotations, making sure that they are made correctly.
            //
            //foreach (var desc in root.DescendantsAndSelf())
            //{
            //    Stack<FieldElementTypeInfo> s = desc.Annotation<Stack<FieldElementTypeInfo>>();
            //    if (s != null)
            //    {
            //        Console.WriteLine(desc.Name.LocalName.PadRight(20));
            //        foreach (var item in s)
            //        {
            //            Console.WriteLine("    {0:0000} {1}", item.Id, item.FieldElementType.ToString());
            //            Console.ReadKey();
            //        }
            //    }
            //}

            var cachedAnnotationInformation = new Dictionary<int, List<XElement>>();
            var deletedFields = new DeletedFieldIds();
            foreach (var desc in root.DescendantsTrimmed(d => d.Name == W.rPr || d.Name == W.pPr))
            {
                Stack<FieldElementTypeInfo>? s = desc.Annotation<Stack<FieldElementTypeInfo>>();

                if (s != null )
                {
                    if (desc.Name == W.fldChar && s.TryPeek(out var top) && top.FieldElementType == FieldElementTypeEnum.Begin
                        && desc.Ancestors().Any(a => a.Name == W.del || a.Name == W.moveFrom))
                        deletedFields.Add(top.Id);
                    foreach (var item in s)
                    {
                        if (item.FieldElementType == FieldElementTypeEnum.InstrText)
                        {
                            if (cachedAnnotationInformation.ContainsKey(item.Id))
                            {
                                cachedAnnotationInformation[item.Id].Add(desc);
                            }
                            else
                            {
                                cachedAnnotationInformation.Add(item.Id, new List<XElement>() { desc });
                            }
                        }
                    }
                }
            }
            root.AddAnnotation(cachedAnnotationInformation);
            root.AddAnnotation(deletedFields);
        }

        /// <summary>Whether field <paramref name="id"/>'s begin character is deleted or moved away.</summary>
        internal static bool IsDeletedField(XElement root, int id) =>
            root.Annotation<DeletedFieldIds>()?.Contains(id) == true;

        /// <summary>The ids of the fields whose begin character is deleted or moved away.</summary>
        private sealed class DeletedFieldIds : HashSet<int>
        {
        }

        private enum State
        {
            InToken,
            InWhiteSpace,
            InQuotedToken,
            OnOpeningQuote,
            OnClosingQuote,
            OnBackslash,
        }

        private static string[] GetTokens(string field)
        {
            State state = State.InWhiteSpace;
            int tokenStart = 0;
            char quoteStart = char.MinValue;
            List<string> tokens = new List<string>();
            for (int c = 0; c < field.Length; c++)
            {
                if (Char.IsWhiteSpace(field[c]))
                {
                    if (state == State.InToken)
                    {
                        tokens.Add(field.Substring(tokenStart, c - tokenStart));
                        state = State.InWhiteSpace;
                        continue;
                    }
                    if (state == State.OnOpeningQuote)
                    {
                        tokenStart = c;
                        state = State.InQuotedToken;
                    }
                    if (state == State.OnClosingQuote)
                        state = State.InWhiteSpace;
                    continue;
                }
                if (field[c] == '\\')
                {
                    if (state == State.InQuotedToken)
                    {
                        state = State.OnBackslash;
                        continue;
                    }
                }
                if (state == State.OnBackslash)
                {
                    state = State.InQuotedToken;
                    continue;
                }
                if (field[c] == '"' || field[c] == '\'' || field[c] == 0x201d)
                {
                    if (state == State.InWhiteSpace)
                    {
                        quoteStart = field[c];
                        state = State.OnOpeningQuote;
                        continue;
                    }
                    if (state == State.InQuotedToken)
                    {
                        if (field[c] == quoteStart)
                        {
                            tokens.Add(field.Substring(tokenStart, c - tokenStart));
                            state = State.OnClosingQuote;
                            continue;
                        }
                        continue;
                    }
                    if (state == State.OnOpeningQuote)
                    {
                        if (field[c] == quoteStart)
                        {
                            state = State.OnClosingQuote;
                            continue;
                        }
                        else
                        {
                            tokenStart = c;
                            state = State.InQuotedToken;
                            continue;
                        }
                    }
                    continue;
                }
                if (state == State.InWhiteSpace)
                {
                    tokenStart = c;
                    state = State.InToken;
                    continue;
                }
                if (state == State.OnOpeningQuote)
                {
                    tokenStart = c;
                    state = State.InQuotedToken;
                    continue;
                }
                if (state == State.OnClosingQuote)
                {
                    tokenStart = c;
                    state = State.InToken;
                    continue;
                }
            }
            if (state == State.InToken)
                tokens.Add(field.Substring(tokenStart, field.Length - tokenStart));
            return tokens.ToArray();
        }

        public static FieldInfo ParseField(string field)
        {
            FieldInfo emptyField = new FieldInfo
            {
                FieldType = "",
                Arguments = new string[] { },
                Switches = new string[] { },
            };

            if (field.Length == 0)
                return emptyField;
            string? fieldType = field.TrimStart().Split(' ').FirstOrDefault();
            if (fieldType == null)
                return emptyField;
            if (fieldType.ToUpper() != "HYPERLINK" &&
                fieldType.ToUpper() != "REF" &&
                fieldType.ToUpper() != "SEQ" &&
                fieldType.ToUpper() != "STYLEREF" &&
                fieldType.ToUpper() != "LISTNUM" &&
                fieldType.ToUpper() != "PAGE")
                return emptyField;
            string[] tokens = GetTokens(field);
            if (tokens.Length == 0)
                return emptyField;
            FieldInfo fieldInfo = new FieldInfo()
            {
                FieldType = tokens[0],
                Switches = tokens.Where(t => t[0] == '\\').ToArray(),
                Arguments = tokens.Skip(1).Where(t => t[0] != '\\').ToArray(),
            };
            return fieldInfo;
        }

        public class FieldInfo
        {
            public string FieldType = "";
            public string[] Switches = Array.Empty<string>();
            public string[] Arguments = Array.Empty<string>();
        }

        public enum FieldElementTypeEnum
        {
            Begin,
            InstrText,
            Separate,
            Result,
            End,
        };

        public class FieldElementTypeInfo
        {
            public int Id;
            public FieldElementTypeEnum FieldElementType;
        }

        public class FieldElementTypeStack
        {
            public int Id;
            public Stack<FieldElementTypeInfo>? FiStack;
        }
    }
}
