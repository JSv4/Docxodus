#!/usr/bin/env python3
"""Private Writer instance for the real-clipboard Doom recording (requires python3-uno).

Commands arrive as JSON lines. Paste dispatches Writer's native unformatted-paste
command; document text is never injected through UNO. Responses inspect Writer's
own text, graphics and saved ODT, independently of the browser producer.
"""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue


def prop(name, value):
    result = PropertyValue()
    result.Name, result.Value = name, value
    return result


scratch = Path(sys.argv[1]).resolve()
port = int(sys.argv[2])
log = (scratch / 'libreoffice.log').open('w')
office = subprocess.Popen([
    os.environ.get('LIBREOFFICE', 'libreoffice'),
    f'-env:UserInstallation={(scratch / "lo-profile").as_uri()}',
    '--nologo', '--nodefault', '--nofirststartwizard', '--norestore',
    f'--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ServiceManager',
], stdout=log, stderr=log)
doc = None
try:
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    deadline = time.monotonic() + 30
    while True:
        try:
            context = resolver.resolve(
                f'uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext')
            break
        except Exception:
            if time.monotonic() >= deadline or office.poll() is not None:
                raise RuntimeError(f'Writer did not start; see {scratch / "libreoffice.log"}')
            time.sleep(.2)
    services = context.ServiceManager
    desktop = services.createInstanceWithContext('com.sun.star.frame.Desktop', context)
    doc = desktop.loadComponentFromURL('private:factory/swriter', '_blank', 0, ())
    controller = doc.getCurrentController()
    window = controller.getFrame().getContainerWindow()
    window.setPosSize(0, 0, 1100, 1050, 15)
    dispatcher = services.createInstanceWithContext('com.sun.star.frame.DispatchHelper', context)

    def dispatch(command, args=()):
        dispatcher.executeDispatch(controller.getFrame(), command, '', 0, args)

    def focus():
        window.setVisible(True)
        window.toFront()
        window.setFocus()

    # Format the empty destination. All content will arrive via the OS clipboard.
    page = doc.StyleFamilies.getByName('PageStyles').getByName(controller.getViewCursor().PageStyleName)
    page.LeftMargin = page.RightMargin = 1200
    page.TopMargin = page.BottomMargin = 1200
    cursor = controller.getViewCursor()
    cursor.CharFontName = 'Courier New'
    cursor.CharHeight = 2
    cursor.CharWeight = 150
    cursor.CharColor = 0xFFFFFF
    cursor.CharKerning = 7  # 0.2 pt, in hundredths of a millimeter
    cursor.ParaBackColor = 0x000000
    cursor.ParaTopMargin = cursor.ParaBottomMargin = 0
    pitch = uno.createUnoStruct('com.sun.star.style.LineSpacing')
    pitch.Mode, pitch.Height = 3, 60  # FIX, 1.7 pt
    cursor.ParaLineSpacing = pitch
    view = controller.getViewSettings()
    view.ZoomType, view.ZoomValue = 3, 125  # BY_VALUE
    print(json.dumps({'ready': True}), flush=True)

    for line in sys.stdin:
        try:
            request = json.loads(line)
            action = request['action']
            result = {}
            if action == 'focus':
                focus()
            elif action == 'paste':
                focus()
                dispatch('.uno:PasteUnformatted')
                deadline = time.monotonic() + 10
                while not doc.Text.String and time.monotonic() < deadline:
                    time.sleep(.05)
                text = doc.Text.String
                expected = Path(request['expected']).read_text()
                result = {
                    'matchesClipboardSource': text == expected,
                    'printableCharacters': sum(32 <= ord(c) <= 126 for c in text),
                    'lineBreaks': text.count('\n'),
                    'nonAsciiCharacters': sum(not (32 <= ord(c) <= 126 or c == '\n') for c in text),
                    'graphicObjects': doc.GraphicObjects.Count,
                    'textFrames': doc.TextFrames.Count,
                    'sha256': hashlib.sha256(text.encode('ascii')).hexdigest(),
                }
                if not result['matchesClipboardSource'] or result['nonAsciiCharacters'] or result['graphicObjects']:
                    raise RuntimeError(f'Clipboard proof failed: {result}')
                # Move the actual caret to the top so the full picture is visible.
                controller.getViewCursor().gotoStart(False)
                view.ZoomType, view.ZoomValue = 3, 125
                doc.storeAsURL(Path(request['odt']).resolve().as_uri(),
                               (prop('FilterName', 'writer8'), prop('Overwrite', True)))
                Path(request['text']).write_text(text, encoding='ascii')
            elif action == 'select':
                # Select a real HUD substring, without changing any characters.
                cursor = controller.getViewCursor()
                cursor.gotoStart(False)
                offset = request['offset']
                while offset:
                    step = min(offset, 30000)
                    cursor.goRight(step, False)
                    offset -= step
                cursor.goRight(request['length'], True)
                result = {'selectedText': cursor.String}
            elif action == 'enlarge':
                cursor = controller.getViewCursor()
                before = doc.Text.String
                dispatch('.uno:FontHeight', (prop('FontHeight.Height', 12.0),))
                pitch = uno.createUnoStruct('com.sun.star.style.LineSpacing')
                pitch.Mode, pitch.Height = 0, 110  # proportional, so enlarged glyphs fit
                cursor.ParaLineSpacing = pitch
                cursor.CharKerning = 0
                result = {'charactersUnchanged': doc.Text.String == before, 'fontPoints': cursor.CharHeight}
                cursor.collapseToStart()
                cursor.goRight(1, False)  # keep the toolbar inside the enlarged text
            elif action == 'screenshot':
                from PIL import ImageGrab
                ImageGrab.grab(xdisplay=os.environ['DISPLAY']).save(request['path'])
            elif action == 'quit':
                print(json.dumps({'ok': True}), flush=True)
                break
            else:
                raise ValueError(f'Unknown action: {action}')
            print(json.dumps({'ok': True, **result}), flush=True)
        except Exception as error:
            print(json.dumps({'error': str(error)}), flush=True)
finally:
    if doc is not None:
        doc.setModified(False)
        doc.close(True)
    office.terminate()
    try:
        office.wait(timeout=10)
    except subprocess.TimeoutExpired:
        office.kill()
        office.wait()
    log.close()
