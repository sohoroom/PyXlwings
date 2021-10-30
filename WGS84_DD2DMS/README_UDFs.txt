One-time Excel preparations

Enable Trust access to the VBA project object model under File > Options > Trust Center > Trust Center Settings > Macro Settings.
You only need to do this once. Also, this is only required for importing the functions, i.e. end users won’t need to bother about this.

Install the add-in via command prompt: xlwings addin install (see Add-in & Settings).

#comtypes pywin32是使用xlwings需要依赖的包。需要额外安装