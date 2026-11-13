# The COM add-in front-end for xlOil's settings ribbon. The callbacks live in
# xloil.ribbon_core, shared with the .xlam front-end (xloil.xlam_ribbon).
import xloil
from xloil import ribbon_core
from xloil.ribbon_core import *            # noqa: F401,F403  (callbacks looked up by name)
from xloil.ribbon_core import _features_menu_xml

@xloil.func
def test_name(x):
    return x

@xloil.func
async def test_name2(x):
    yield x


def _ribbon_func_map(func: str):
    # Just finds the function with the given name in this module 
    xloil.log.debug(f"Calling xlOil Ribbon '{func}'...")
    return getattr(ribbon_core, func, None) or globals()[func]

_ribbon_ui = xloil.ExcelGUI(ribbon=r'''
    <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
      <ribbon>
        <tabs>
          <tab id="tab0" label="xlOil Py">
            <group id="grp0" autoScale="false" centerVertically="false" label="Environment">
              <dropDown id="ddc0" label="Environment" 
                screentip="Python environments listed in the Windows Registry"
                supertip="Environment changes only take effect after restarting Excel"
                onAction="set_python_environment"
                getItemCount="get_python_environment_count"
                getItemLabel="get_python_environment"
                getSelectedItemIndex="get_python_environment_selected" />
              <editBox id="PYTHONEXECUTABLE" label="PYTHONEXECUTABLE" 
                screentip="The location of python.exe in the distribution"
                supertip="Environment changes only take effect after restarting Excel"
                sizeString="c:/a/path/is/this/size"
                getText="get_python_home" 
                onChange="set_python_home" />
              <editBox id="PYTHONPATH" label="PYTHONPATH" sizeString="c:/a/path/is/this/size"
                screentip="A semi-colon separated list of module search directories"
                supertip="Prefer to use this for system paths and add user directories to the Search Paths"
                getText="get_python_path" 
                onChange="set_python_path" />
            </group>
            <group id="grp1" autoScale="false" centerVertically="false" label="Modules">
              <!--FEATURES-->
              <editBox id="ebx3" label="Load Modules" sizeString="a module; another module; another"
                screentip="Python modules to load"
                supertip="The modules must be available on python's sys.path. Use a comma to separate."
                getText="get_load_modules" 
                onChange="set_load_modules"/>
              <editBox id="ebx4" label="Search Paths" sizeString="a module; another module; another"
                screentip="Paths added to python's sys.path"
                supertip="Prefer to add user directories here rather than editing PYTHONPATH directly. Use the usual semi-colon (;) separator and the path separator (\)"
                getText="get_user_search_path"
                onChange="set_user_search_path"/>
              <labelControl id="RestartLabel" label="!Restart Required!" 
                getVisible="get_restart_label_visible" />
            </group>
            <group id="grpTools" autoScale="false" centerVertically="false" label="Tools">
              <button id="btnDoctor" size="large" label="Doctor" imageMso="ReviewProtectDocument"
                onAction="press_doctor"
                screentip="Check the installation"
                supertip="Checks Excel/Python bitness, the C++ runtime, required packages, the add-in and settings, and says how to fix any problem"/>
              <button id="btnReload" size="large" label="Reload" imageMso="Refresh"
                onAction="press_reload"
                screentip="Reload Python modules"
                supertip="Reloads every module xlOil registered functions from, without restarting Excel"/>
              <button id="btnSettings" size="large" label="Settings File" imageMso="FileProperties"
                onAction="press_edit_settings"
                screentip="Open xlOil.ini"
                supertip="Opens the settings file in your text editor. Most changes need an Excel restart"/>
            </group>
            <group id="grp2" autoScale="false" centerVertically="false" label="Debugging" >
              <button id="btn5" size="large" label="Console" imageMso="ContainerGallery" 
                onAction="press_open_console"/>
              <button id="qtconsole" size="large" label="QtConsole" imageMso="ContainerGallery" 
                onAction="press_open_qtconsole"
                screentip="Opens a Jupyter console to an inprocess kernel"
                supertip="The console is connected to the Excel application, so can register functions and manipulate workbooks"/>
              <button id="btn6" size="large" label="Open Log" imageMso="ZoomCurrent75" 
                onAction="press_open_log"
                screentip="Opens the log file"
                supertip="Uses the editor configured to open '.log' files"/>
              <dropDown id="ddc8" label="Log Level" 
                screentip="The threshold level to control which messages are written to the log file"
                getSelectedItemIndex="get_log_level_selected" 
                getItemCount="get_log_level_count"
                getItemLabel="get_log_level"
                onAction="set_log_level" >
              </dropDown>
              <dropDown id="ddc9" label="Debugger" 
                screentip="Debugger used"
                supertip="Check the documentation for details on debugging"
                getSelectedItemIndex="get_debugger_selected" 
                getItemCount="get_debugger_count"
                getItemLabel="get_debugger"
                onAction="set_debugger" >
              </dropDown>
              <editBox id="ebxDebugPy" label="DebugPy Port"
                screentip="Connection port used if the debugpy/vscode debugger is selected"
                getText="get_debugpy_port"
                onChange="set_debugpy_port"/>
            </group>
            <group id="grp3" autoScale="false" centerVertically="false" label="Other" >
              <editBox id="ebxDateFormats" label="Date Formats" 
                screentip="Date formats tried when parsing strings as dates"
                supertip="Uses 'std::get_time' formats: %Y year, %m month number, %b month name, etc." 
                getText="get_date_formats" 
                onChange="set_date_formats"/>
              <checkBox id="cbxPropagation" label="Propagate Errors"
                screentip="If enabled, any error code (eg. \#NUM!) arguments are passed through as the function's return 
                            value, otherwise all argument values are handled by the function"
                supertip="(requires restart)"
                getPressed="get_error_propagation"              
                onAction="set_error_propagation"/>
              <button id="fixNameErrors" size="normal" label="Fix #NAME!" imageMso="ErrorChecking" 
                onAction="_fix_name_errors"
                screentip="Marks #NAME! errors in current workbook for recalculation."
                supertip="These errors cannot simply be resolved just by Ctrl-Alt-F9. May not be performant in 
                          large workbooks with many errors."/>
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>
    '''.replace("<!--FEATURES-->", _features_menu_xml()),
    funcmap=_ribbon_func_map)

def _com_invalidate(control_id=None):
    if control_id:
        _ribbon_ui.invalidate(control_id)
    else:
        _ribbon_ui.invalidate()

ribbon_core._invalidate = _com_invalidate
