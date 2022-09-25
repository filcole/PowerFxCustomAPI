using Microsoft.Xrm.Sdk;
using System;
using Microsoft.PowerFx;
//using Microsoft.PowerFx.Types;
//using Microsoft.PowerFx.Core;
//using Microsoft.PowerFx;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using pgc.EarlyBindings;
using Microsoft.PowerFx.Types;
using YamlDotNet.Core;
using System.Xml;
using Newtonsoft.Json;
using System.Text.RegularExpressions;
using YamlDotNet.RepresentationModel;
using System.IO;
using Microsoft.PowerFx.Core;

namespace Plugin
{
    /// <summary>
    /// Plugin development guide: https://docs.microsoft.com/powerapps/developer/common-data-service/plug-ins
    /// Best practices and guidance: https://docs.microsoft.com/powerapps/developer/common-data-service/best-practices/business-logic/
    /// </summary>
    public class PowerFxEvaluate : PluginBase
    {
        private const string OptionFormatTable = "FormatTable";

        internal class Formula
        {
            internal string Name { get; set; }
            internal string Expression { get; set; }
        }

        public PowerFxEvaluate(string unsecureConfiguration, string secureConfiguration)
            : base(typeof(PowerFxEvaluate))
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

            var request = new pgc_PowerFxEvaluateRequest
            {
                Parameters = localPluginContext.PluginExecutionContext.InputParameters
            };

            localPluginContext.Trace($"23 Context: {EscapeJSON(request.Context)}");
            localPluginContext.Trace($"Yaml: {request.Yaml}");

            // Message: The type initializer for 'Microsoft.PowerFx.Core.Types.Enums.EnumStoreBuilder' threw an exception.
            // System.Resources.MissingSatelliteAssemblyException: The satellite assembly named "Microsoft.PowerFx.Core.resources.dll, PublicKeyToken=31bf3856ad364e35" for fallback culture "en-US" either could not be found or could not be loaded. This is generally a setup problem. Please consider reinstalling or repairing the application.

            try
            {
                localPluginContext.Trace("Starting recalc engine");
                var engine = new RecalcEngine();

                var fxContext = String.IsNullOrWhiteSpace(request.Context) ? "{}" : request.Context;

                localPluginContext.Trace("Setting input");

                // We may be passed a JSON context, but if it's not passed then create an empty object.
                var input = (RecordValue)FormulaValue.FromJson(fxContext);

                localPluginContext.Trace("Starting reading of YAML formulae");

                // Try and get the list of formuale from the passed Yaml
                List<Formula> formulae;
                try
                {
                    // Read the Yaml and parse into a list of variables and expressions
                    formulae = GetFormulae(request.Yaml, localPluginContext);
                }
                catch (YamlException ex)
                {
                    var errmsg = $"Exception {ex.Message} extracting formula from YAML. Inner exception: {ex.InnerException}";
                    localPluginContext.Trace(errmsg);
                    throw new InvalidPluginExecutionException(errmsg);
                }

                localPluginContext.Trace($"Processing {formulae.Count} formulae");

                // Evaulate each formula in turn, store the result of each formula back in the PowerFx engine
                // so that it can be used by later formulas.
                foreach (var f in formulae)
                {
                    try
                    {
                        engine.UpdateVariable(f.Name, engine.Eval(f.Expression, input));
                    }
                    catch (Exception ex)
                    {
                        localPluginContext.Trace(String.Format("Exception: {0} on formula '{1}'", ex.Message, f.Expression));
                        throw new InvalidPluginExecutionException($"PowerFx error on forumla '{f.Expression}': {ex.Message}");
                    }
                }

                // Note: Integers serialise as decimal numbers, but the Parse Json step in Power Automate will
                // happily converts them back to integers within Power Automate.
                var output = new Dictionary<string, Object>();
                foreach (var f in formulae)
                {
                    // Yaml expression may contain a variable multiple times,
                    // but it only needs to be returned once.
                    if (!output.ContainsKey(f.Name))
                    {
                        output[f.Name] = engine.GetValue(f.Name).ToObject();
                    }
                }

                // Format the output so that it's easier to see in Power Automate.
                string json = JsonConvert.SerializeObject(output); //, Newtonsoft.Json.Formatting.Indented);
                localPluginContext.Trace("Successful response");

                var response = new pgc_PowerFxEvaluateResponse
                {
                    Results = localPluginContext.PluginExecutionContext.OutputParameters
                };
                response.Output = json;
            }
            catch (Exception e)
            {
                localPluginContext.Trace($"Got Exception");
                localPluginContext.Trace($"Message: {e.Message}");
                localPluginContext.Trace($"InnerException: {e.InnerException}");
            }
        }

        // Escape JSON so that it can be passed to tracing service which used String.Format()
        private object EscapeJSON(string context)
        {
            if (context == null)
            {
                return null;
            }
            return context.Replace("{", "{{").Replace("}", "}}");
        }

        // Build a list of forumlae from the Yaml that's passed to the function
        private List<Formula> GetFormulae(string formulaYaml, ILocalPluginContext localPluginContext)
        {
            // Read and parse the Yaml
            var yaml = new YamlStream();
            yaml.Load(new StringReader(formulaYaml));

            var formulae = new List<Formula>();

            // Fetch all nodes, it's simpler than trying to navigate the tree structure.
            // There's room to improve this!
            foreach (var node in yaml.Documents[0].AllNodes)
            {
                // We're only interested in the mapping nodes, but these may be top-level, or at the bottom
                // of the tree structure that is Yaml SequenceNodes/MappingNodes.
                if (node is YamlMappingNode mapping)
                {
                    foreach (var entry in mapping.Children)
                    {
                        if (entry.Value is YamlScalarNode val)
                        {
                            var expression = RemoveComments(val.Value).Trim();

                            localPluginContext.Trace($"expression: {expression}");

                            if (expression.StartsWith("="))
                            {
                                var name = ((YamlScalarNode)entry.Key).Value;
                                formulae.Add(new Formula
                                {
                                    Name = name,
                                    // Remove the first character (=)
                                    Expression = expression.Substring(1),
                                });
                            }
                        }
                    }
                }
            }
            return formulae;
        }

        // Remove single line comments '//' and multi-line comments /* xxx */
        // Thank you https://stackoverflow.com/a/3524689
        private static string RemoveComments(string input)
        {
            var blockComments = @"/\*(.*?)\*/";
            var lineComments = @"//(.*?)\r?\n";
            var strings = @"""((\\[^\n]|[^""\n])*)""";
            var verbatimStrings = @"@(""[^""]*"")+";

            return Regex.Replace(input,
                blockComments + "|" + lineComments + "|" + strings + "|" + verbatimStrings,
                me =>
                {
                    if (me.Value.StartsWith("/*") || me.Value.StartsWith("//"))
                    {
                        return me.Value.StartsWith("//") ? Environment.NewLine : "";
                    }
                    // Keep the literal strings
                    return me.Value;
                },
                RegexOptions.Singleline
            );
        }
    }
}
