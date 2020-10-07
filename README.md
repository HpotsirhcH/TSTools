# TSTools

Tasksequence Tools

Some Tools and Dlls for SCCM/MDT Tasksequences


# TSTools.TSEnv
This DLL interact with the Tasksequnce 
### Implemented Features

Get Set TS Variable
#### Usage:

<code>
using TSTools.TSEnv;
private static string GETTSValue(string tsvalue) => TSEnv.GetTSVariable(tsvalue);
</code>

Remove Progress Window and Open it again
#### Usage:
<code>
using TSTools.TSEnv;
try { TSEnv.CloseTSProgressUI(); } catch { } // close the Progress Gui
try { TSEnv.ShowTSProgressUI(); } catch { } // Show it again.
</code>