using System;

namespace TSTools.TSEnv
{
    public class TSEnv
    {
        /// <summary>
        /// Get the Value of a TS Variable
        /// </summary>
        /// <param name="TSVarName">What TS Variable should be read?</param>
        /// <returns>string</returns>
        public static string GetTSVariable(string TSVarName)
        {
            string returnvalue;
            // Binding Of Microsoft.SMS.TSEnvironment COM Object.
            System.Type objType = System.Type.GetTypeFromProgID("Microsoft.SMS.TSEnvironment");
            dynamic smsObject = System.Activator.CreateInstance(objType);

            returnvalue = smsObject.Value[TSVarName];

            // Release Microsoft.SMS.TSEnvironment object.
            if (System.Runtime.InteropServices.Marshal.IsComObject(smsObject) == true)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(smsObject);
            }

            return returnvalue;
        }

        /// <summary>
        /// Set a TS Variable
        /// </summary>
        /// <param name="TSVarName">TS Variable name</param>
        /// <param name="Value">Value of this Variable</param>
        public static void SetTSVariable(string TSVarName, string Value)
        {
            // Binding Of Microsoft.SMS.TSEnvironment COM Object.
            System.Type objType = System.Type.GetTypeFromProgID("Microsoft.SMS.TSEnvironment");
            dynamic tsEnv = System.Activator.CreateInstance(objType);

            tsEnv.Value[TSVarName] = Value;

            // Release Microsoft.SMS.TSEnvironment object.
            if (System.Runtime.InteropServices.Marshal.IsComObject(tsEnv) == true)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tsEnv);
            }
        }

        /// <summary>
        /// Close the TS Progress UI for not overlay the Program.
        /// </summary>
        public static void CloseTSProgressUI()
        {
            Type tsProgress = Type.GetTypeFromProgID("Microsoft.SMS.TSProgressUI");
            dynamic comObject = Activator.CreateInstance(tsProgress);

            comObject.CloseProgressDialog();

            if (System.Runtime.InteropServices.Marshal.IsComObject(comObject) == true)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(comObject);
        }

        /// <summary>
        /// Show the TS Progress UI.
        /// </summary>
        public static void ShowTSProgressUI()
        {
            Type objType = Type.GetTypeFromProgID("Microsoft.SMS.TSEnvironment");
            dynamic tsEnv = Activator.CreateInstance(objType);

            Type progressType = Type.GetTypeFromProgID("Microsoft.SMS.TSProgressUI");
            dynamic tsProgress = Activator.CreateInstance(progressType);

            string orgName = tsEnv.Value["_SMSTSOrgName"];
            string pkgName = tsEnv.Value["_SMSTSPackageName"];
            string customMessage = tsEnv.Value["_SMSTSCustomProgressDialogMessage"];
            string currentAction = tsEnv.Value["_SMSTSCurrentActionName"];
            string nextInstructor = tsEnv.Value["_SMSTSNextInstructionPointer"];
            string maxStep = tsEnv.Value["_SMSTSInstructionTableSize"];

            tsProgress.ShowTSProgress(orgName, pkgName, customMessage, currentAction, nextInstructor, maxStep);

            if (System.Runtime.InteropServices.Marshal.IsComObject(tsEnv) == true)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tsEnv);

            if (System.Runtime.InteropServices.Marshal.IsComObject(tsProgress) == true)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tsProgress);
        }
    }
}
