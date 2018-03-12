// Copyright (c) Benjamin Trent. All rights reserved. See LICENSE file in project root

namespace PMCCommand
{
    using System;
    using System.Globalization;
    using System.IO;
    using EnvDTE;
    using EnvDTE80;

    /// <summary>
    /// This is the least hacky way I could find for running commands in the PMC from the command line
    /// This program will open an instance of VisualStudio and execute the passed in commands into the PMC directly
    /// From all that I could find, there was nothing that allowed this without passing through VisualStudio, which is depressing
    /// </summary>
    public class Program
    {
        private const string CmdNameForPMC = "View.PackageManagerConsole";
        private static bool retry = true;
        private static ExecutionState state = ExecutionState.NOT_STARTED;
        private static object stateMutex = new object();

        private enum ExecutionState
        {
            NOT_STARTED,
            VS_OPENED,
            PROJECT_OPENED,
            NUGET_OPENED
        }

        private static string VSVersion { get; set; }

        private static string ProjectPath { get; set; }

        private static string NuGetCmd { get; set; }

        private static bool Debug { get; set; }

        private static DTE DTE { get; set; }

        private static string LockFile { get; set; }

        private static string NuGetOutputFile { get; set; }

        private static void CleanUp()
        {
            if (DTE != null)
            {
                if (DTE.Solution != null)
                {
                    DTE.Solution.Close(true);
                }

                DTE.Quit();
            }
        }

        /// <summary>
        /// STAThread is necessary for the MessageFilter
        /// </summary>
        /// <param name="args"> Command line args</param>
        [STAThread]
        private static void Main(string[] args)
        {
            var options = new CmdLineOptions();
            if (CommandLine.Parser.Default.ParseArgumentsStrict(args, options))
            {
                ProjectPath = options.ProjectPath;
                NuGetCmd = options.NuGetCommand;
                Debug = options.Debug;

                if (string.IsNullOrWhiteSpace(NuGetCmd))
                {
                    throw new Exception("nugetcommand parameter cannot be empty.");
                }

                if (string.IsNullOrWhiteSpace(ProjectPath))
                {
                    throw new Exception("project parameter cannot be empty.");
                }

                VSVersion = "15.0";

                if (!string.IsNullOrWhiteSpace(options.VisualStudioVersion))
                {
                    VSVersion = options.VisualStudioVersion;
                }

                if (!File.Exists(ProjectPath))
                {
                    throw new FileNotFoundException(string.Format("{0} was not found.", ProjectPath));
                }

                Execute();
            }
        }

        /// <summary>
        /// When your external, multi-threaded application calls into Visual Studio, it goes through a COM interface.
        /// COM sometimes has problems dealing properly with threads, especially with respect to timing.
        /// As a result, occasionally the incoming thread from the external application cannot be handled by Visual Studio at the very moment it arrives, resulting in the previously mentioned errors.
        /// This does not occur, however, if you are calling from an application that is running inside Visual Studio (in-proc), such as a macro or an add-in.
        /// For a more detailed explanation about the reasons behind this, see https://msdn.microsoft.com/en-us/library/8sesy69e.aspx.
        /// To avoid these errors, implement an IOleMessageFilter handler function in your application.When you do this,
        ///  if your external application thread calls into Visual Studio and is rejected (that is, it returns SERVERCALL_RETRYLATER from the IOleMessageFilter.HandleIncomingCall method),
        ///  then your application can handle it and either retry or cancel the call.
        /// To do this, initiate the new thread from your Visual Studio application in a single-threaded apartment(STA) and surround your automation code with the IOleMessageFilter handler.
        /// </summary>
        private static void Execute()
        {
            // This is the only way I could find for making sure that the PMC command completed
            LockFile = Path.GetTempFileName();
            using (StreamWriter sw = new StreamWriter(LockFile))
            {
                sw.WriteLine(true);
            }

            NuGetOutputFile = Path.GetTempFileName();
            DTE = GetDTE2();
            MessageFilter.Register();

            Console.CancelKeyPress += (object sender, ConsoleCancelEventArgs e) =>
            {
                DTE.Quit();

                // turn off the IOleMessageFilter
                MessageFilter.Revoke();
            };

            SetDelegatesForDTE();
            DTE.MainWindow.Activate();

            SpinWait(ExecutionState.VS_OPENED);
            DTE.ExecuteCommand("File.OpenProject", ProjectPath);

            SpinWait(ExecutionState.PROJECT_OPENED);
            DTE.ExecuteCommand(CmdNameForPMC);

            SpinWait(ExecutionState.NUGET_OPENED);
            var cmd = string.Format("{0}; $error > {1} ; \"False\" > {2}", NuGetCmd, NuGetOutputFile, LockFile);
            DTE.ExecuteCommand(CmdNameForPMC, cmd);

            var stillRunning = true;
            while (stillRunning)
            {
                System.Threading.Thread.Sleep(500);
                using (StreamReader sr = new StreamReader(LockFile))
                {
                    stillRunning = Convert.ToBoolean(sr.ReadToEnd());
                }
            }

            Console.WriteLine("Completed");
            Console.WriteLine(File.ReadAllText(NuGetOutputFile));
            CleanUp();

            // turn off the IOleMessageFilter
            MessageFilter.Revoke();
        }

        /// <summary>
        /// We need to insure that particular parts of the main STAThread do not continue until feed back of other actions is received
        /// Since we could potentially get a failure from the MessageFilter and a retry.
        /// So, using a Monitor would not work as potentially the `Pulse` would occur before the `Wait`
        /// Hence, the use of a naive state machine
        /// </summary>
        /// <param name="expectedState">The state desired necessary to continue.</param>
        private static void SpinWait(ExecutionState expectedState)
        {
            while (state < expectedState)
            {
                System.Threading.Thread.Sleep(1000);
            }
        }

        /// <summary>
        /// Failures can occure when accessing the newly created DTE2 instance.
        /// This is due to COM Interop errors in Windows.
        /// The MessageFilter should catch failures, but if a particular one leaks through, restart this initial setting
        /// </summary>
        private static void SetDelegatesForDTE()
        {
            try
            {
                DTE.Events.SolutionEvents.Opened += SolutionEvents_Opened;
                DTE.Events.CommandEvents.AfterExecute += CommandEvents_AfterExecute;
                DTE.Events.CommandEvents.BeforeExecute += CommandEvents_BeforeExecute;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Console.WriteLine("Exception encountered: " + ex.Message);
                if (retry)
                {
                    retry = false;
                    System.Threading.Thread.Sleep(500);
                    SetDelegatesForDTE();
                }
            }
        }

        private static void CommandEvents_BeforeExecute(string guid, int id, object customIn, object customOut, ref bool cancelDefault)
        {
            PrintDebugLog(string.Format("Command Sent: GUID: {0}; ID: {1}; CustomIn: {2}; CustomOut: {3}", guid, id, customIn, customOut));
        }

        private static void CommandEvents_AfterExecute(string guid, int id, object customIn, object customOut)
        {
            PrintDebugLog(string.Format("Command Executed: GUID: {0}; ID: {1}; CustomIn: {2}; CustomOut: {3}", guid, id, customIn, customOut));

            // This means that PMC has loaded and loaded its sources
            if (guid == GuidsAndIds.GuidNuGetConsoleCmdSet && id == GuidsAndIds.CmdidNuGetSources)
            {
                TransitionState(ExecutionState.PROJECT_OPENED, ExecutionState.NUGET_OPENED);
            }
            else
            {
                TransitionState(ExecutionState.NOT_STARTED, ExecutionState.VS_OPENED);
            }
        }

        private static void SolutionEvents_Opened()
        {
            TransitionState(ExecutionState.VS_OPENED, ExecutionState.PROJECT_OPENED);
        }

        private static void TransitionState(ExecutionState expectedOldState, ExecutionState newState)
        {
            lock (stateMutex)
            {
                if (state == expectedOldState)
                {
                    state = newState;
                    PrintDebugLog(string.Format("Transitioned from {0} to {1}", expectedOldState, newState));
                }
            }
        }

        private static string GetValue(string key, string defaultValue)
        {
            var result = Environment.GetEnvironmentVariable(key);
            if (string.IsNullOrEmpty(result))
            {
                result = defaultValue;
            }

            return result;
        }

        private static DTE GetDTE2()
        {
            // Get the ProgID for DTE 14.0.
            Type t = Type.GetTypeFromProgID(
                "VisualStudio.DTE." + VSVersion, true);

            // Create a new instance of the IDE.
            object obj = Activator.CreateInstance(t, true);

            // Cast the instance to DTE2 and assign to variable dte.
            DTE2 dte2 = (DTE2)obj;

            // We want to make sure that the devenv is killed when we quit();
            dte2.UserControl = false;
            return dte2.DTE;
        }

        private static void PrintDebugLog(string message)
        {
            if (Debug)
            {
                Console.WriteLine("{0}: {1}", DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.ffff", CultureInfo.InvariantCulture), message);
            }
        }
    }
}
