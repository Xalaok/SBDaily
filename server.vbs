Dim ws
Set ws = Wscript.CreateObject("Wscript.Shell")
ws.run "hugo server --navigateToChanged --disableFastRender --renderToMemory"
Wscript.quit