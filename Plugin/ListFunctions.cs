using Microsoft.PowerFx;
using Microsoft.PowerFx.Core;
using pgc.EarlyBindings;
using System;
using System.Collections.Generic;
using System.Linq;

namespace pgc.PowerFxCustomAPI
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

        // Entry point for custom business logic execution
        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            var version = typeof(RecalcEngine).Assembly.GetName().Version.ToString();
            localPluginContext.Trace($"Microsoft Power Fx Console Formula REPL, Version {version}");

            localPluginContext.Trace("Starting recalc engine");
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

            var response = new pgc_PowerFxListFunctionsResponse
            {
                Results = localPluginContext.PluginExecutionContext.OutputParameters
            };
            response.Functions = functions;
            response.Version = version;
        }
    }
}