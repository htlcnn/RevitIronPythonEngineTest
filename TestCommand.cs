using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using IronPython;
using IronPython.Compiler;
using IronPython.Runtime.Exceptions;
using Microsoft.Scripting;
using Microsoft.Scripting.Hosting;
using System.Windows.Forms;

namespace RevitIronPythonTest
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    public class TestCommand : IExternalCommand
    {
        public string Message { get; private set; } = null;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var engine = CreateEngine();
                var scope = SetupEnvironment(engine);
                //ScriptScope scope = IronPython.Hosting.Python.CreateModule(engine, "__main__");
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "|*.py";
                if (ofd.ShowDialog() != DialogResult.OK)
                {
                    return Result.Cancelled;
                }

                //var sourcePath = @"D:\HTL\Desktop\ironpython\import279.py";
                var sourcePath = ofd.FileName;
                TaskDialog.Show("file", sourcePath);
                var script = engine.CreateScriptSourceFromFile(sourcePath, Encoding.UTF8, SourceCodeKind.File);

                // setting module to be the main module so __name__ == __main__ is True
                var compiler_options = (PythonCompilerOptions)engine.GetCompilerOptions(scope);
                compiler_options.ModuleName = "__main__";
                compiler_options.Module |= IronPython.Runtime.ModuleOptions.Initialize;

                // Setting up error reporter and compile the script
                var errors = new ErrorReporter();
                var command = script.Compile(compiler_options, errors);
                if (command == null)
                {
                    TaskDialog.Show("error", "fail to compile");
                    // compilation failed, print errors and return
                    Message =
                        string.Join("\r\n", "IronPython Traceback:", string.Join("\r\n", errors.Errors.ToArray()));
                    return Result.Cancelled;
                }

                try
                {
                    script.Execute(scope);
                    return Result.Succeeded;
                }
                catch (SystemExitException)
                {
                    // ok, so the system exited. That was bound to happen...
                    return Result.Succeeded;
                }
                catch (Exception exception)
                {
                    string _dotnet_err_message = exception.ToString();
                    string _ipy_err_messages = engine.GetService<ExceptionOperations>().FormatException(exception);

                    _ipy_err_messages =
                        string.Join("\n", "IronPython Traceback:", _ipy_err_messages.Replace("\r\n", "\n"));
                    _dotnet_err_message =
                        string.Join("\n", "Script Executor Traceback:", _dotnet_err_message.Replace("\r\n", "\n"));

                    Message = _ipy_err_messages + "\n\n" + _dotnet_err_message;

                    return Result.Failed;
                }
                finally
                {
                    TaskDialog.Show("info", Message);
                    engine.Runtime.Shutdown();
                    engine = null;
                }

            }
            catch (Exception ex)
            {
                Message = ex.ToString();
                TaskDialog.Show("error", Message);
                return Result.Failed;
            }
        }
        public ScriptEngine CreateEngine()
        {
            var _fullframe = true;
            var flags = new Dictionary<string, object>();

            // default flags
            flags["LightweightScopes"] = true;

            if (_fullframe)
            {
                flags["Frames"] = true;
                flags["FullFrames"] = true;
            }

            var engine = IronPython.Hosting.Python.CreateEngine(flags);

            return engine;
        }
        public ScriptScope SetupEnvironment(ScriptEngine engine)
        {
            var scope = IronPython.Hosting.Python.CreateModule(engine, "__main__");

            SetupEnvironment(engine, scope);

            return scope;
        }
        public void SetupEnvironment(ScriptEngine engine, ScriptScope scope)
        {
            // add two special variables: __revit__ and __vars__ to be globally visible everywhere:            
            var builtin = IronPython.Hosting.Python.GetBuiltinModule(engine);
            //builtin.SetVariable("__revit__", _revit);

            // add the search paths
            //AddEmbeddedLib(engine);

            // reference RevitAPI and RevitAPIUI
            engine.Runtime.LoadAssembly(typeof(Autodesk.Revit.DB.Document).Assembly);
            engine.Runtime.LoadAssembly(typeof(Autodesk.Revit.UI.TaskDialog).Assembly);

            // also, allow access to the RPL internals
            //engine.Runtime.LoadAssembly(typeof(PyRevitLoader.ScriptExecutor).Assembly);
        }
    }
    public class ErrorReporter : ErrorListener
    {
        public List<String> Errors = new List<string>();

        public override void ErrorReported(ScriptSource source, string message, SourceSpan span, int errorCode, Severity severity)
        {
            Errors.Add(string.Format("{0} (line {1})", message, span.Start.Line));
        }

        public int Count
        {
            get { return Errors.Count; }
        }
    }

}
