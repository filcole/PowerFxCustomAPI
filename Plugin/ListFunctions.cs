using Microsoft.Xrm.Sdk;
using System;
using Microsoft.PowerFx;
using Microsoft.PowerFx.Types;
using Microsoft.PowerFx.Core;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using pgc.EarlyBindings;

namespace Plugin
{
    /// <summary>
    /// Plugin development guide: https://docs.microsoft.com/powerapps/developer/common-data-service/plug-ins
    /// Best practices and guidance: https://docs.microsoft.com/powerapps/developer/common-data-service/best-practices/business-logic/
    /// </summary>
    public class ListFunctions : PluginBase
    {
        private const string OptionFormatTable = "FormatTable";

        public ListFunctions(string unsecureConfiguration, string secureConfiguration)
            : base(typeof(ListFunctions))
        {
            // TODO: Implement your custom configuration handling
            // https://docs.microsoft.com/powerapps/developer/common-data-service/register-plug-in#set-configuration-data
        }

        protected static RecalcEngine GetEngine()
        {
            Features toenable = 0;
            foreach (Features feature in (Features[])Enum.GetValues(typeof(Features)))
                toenable |= feature;

            var config = new PowerFxConfig(toenable);

            var OptionsSet = new OptionSet("Options", DisplayNameUtility.MakeUnique(new Dictionary<string, string>()
                                                {
                                                        { OptionFormatTable, OptionFormatTable },
                                                }
                                           ));

            config.AddOptionSet(OptionsSet);

            return new RecalcEngine(config);
        }

        // Entry point for custom business logic execution
        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            localPluginContext.Trace("My plugin executed 15");

            var version = typeof(RecalcEngine).Assembly.GetName().Version.ToString();
            localPluginContext.Trace($"Microsoft Power Fx Console Formula REPL, Version {version}");

            localPluginContext.Trace("Starting recalc engine");

            CultureInfo.CurrentCulture = new CultureInfo("en-US");

            // Message: The type initializer for 'Microsoft.PowerFx.Core.Types.Enums.EnumStoreBuilder' threw an exception.
            // System.Resources.MissingSatelliteAssemblyException: The satellite assembly named "Microsoft.PowerFx.Core.resources.dll, PublicKeyToken=31bf3856ad364e35" for fallback culture "en-US" either could not be found or could not be loaded. This is generally a setup problem. Please consider reinstalling or repairing the application.

            string[] functions = null;
            try
            {
                var engine = new RecalcEngine();
                functions = engine.GetAllFunctionNames().OrderBy(x => x).ToArray();
                localPluginContext.Trace($"Functions: {String.Join(", ", functions)}");
            }
            catch (Exception e)
            {
                localPluginContext.Trace($"Got Exception");
                localPluginContext.Trace($"Message: {e.Message}");
                localPluginContext.Trace($"InnerException: {e.InnerException}");
            }

            //localPluginContext.PluginExecutionContext.OutputParameters["out"] = "Blaaah";
            //localPluginContext.PluginExecutionContext.OutputParameters["Functions"] = new String[] { "a", "b" };
            var response = new pgc_PowerFxListFunctionsResponse
            {
                Results = localPluginContext.PluginExecutionContext.OutputParameters
            };
            response.Functions = functions;
            

            // TODO: Implement your custom business logic

            // Check for the entity on which the plugin would be registered
            //if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            //{
            //    var entity = (Entity)context.InputParameters["Target"];

            //    // Check for entity name on which this plugin would be registered
            //    if (entity.LogicalName == "account")
            //    {

            //    }myplugin		24/09/2022 04:35:11	FileStore	1.0.0.0		7fd009ec86136436	neutral	Sandbox	False	

            //}
        }
    }
}
