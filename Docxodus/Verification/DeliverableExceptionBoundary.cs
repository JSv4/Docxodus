// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.Verification;

/// <summary>Shared detector boundary that preserves genuinely fatal/control-flow exceptions.</summary>
internal static class DeliverableExceptionBoundary
{
    internal static bool IsRecoverable(Exception exception)
    {
        if (exception is OutOfMemoryException
            or StackOverflowException
            or AccessViolationException
            or AppDomainUnloadedException
            or BadImageFormatException
            or CannotUnloadAppDomainException
            or InvalidProgramException
            or System.Runtime.InteropServices.SEHException
            or System.Threading.ThreadAbortException
            or OperationCanceledException)
            return false;
        if (exception is AggregateException aggregate)
            return aggregate.InnerExceptions.All(IsRecoverable);
        return exception.InnerException is null || IsRecoverable(exception.InnerException);
    }
}
