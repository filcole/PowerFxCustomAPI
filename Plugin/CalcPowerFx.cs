using Microsoft.PowerFx;
using Microsoft.PowerFx.Types;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using pgc.EarlyBindings;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using YamlDotNet.Core;
using YamlDotNet.RepresentationModel;

namespace pgc.PowerFxCustomAPI
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

            // Note: I'm very verbose with logging in the plugin - probably excessively so! Probably best to remove these so that
            // the trace log buffer doesn't get exceeded for large payloads
            localPluginContext.Trace($"Context: {EscapeJSON(request.Context)}");
            localPluginContext.Trace($"Yaml: {EscapeJSON(request.Yaml)}");

            // NOTE: If you get the following error
            //   Message: The type initializer for 'Microsoft.PowerFx.Core.Types.Enums.EnumStoreBuilder' threw an exception.
            //   System.Resources.MissingSatelliteAssemblyException: The satellite assembly named "Microsoft.PowerFx.Core.resources.dll, PublicKeyToken=31bf3856ad364e35" for fallback culture "en-US" either could not be found or could not be loaded. This is generally a setup problem. Please consider reinstalling or repairing the application.
            // Then it's because the Microsoft.PowerFx.Core.resources.dll needs to be copied into the 'Plugin' folder from the nuget cache.  This may be a Power Fx nuget packaging issue.

            try
            {
                localPluginContext.Trace("Starting recalc engine");
                var engine = new RecalcEngine();

                // Handle empty context
                var fxContext = String.IsNullOrWhiteSpace(request.Context) ? "{}" : request.Context;

                // Convert our context into a record value for use when evaluating
                localPluginContext.Trace("Setting input");
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
                    localPluginContext.Trace(EscapeJSON(errmsg));
                    throw new InvalidPluginExecutionException(errmsg);
                }

                localPluginContext.Trace($"Processing {formulae.Count} formulae");

                // Evaulate each formula in turn, store the result of each formula back in the Power Fx engine
                // so that it can be used by later formulas.
                foreach (var f in formulae)
                {
                    try
                    {
                        engine.UpdateVariable(f.Name, engine.Eval(f.Expression, input));
                    }
                    catch (Exception ex)
                    {
                        localPluginContext.Trace(EscapeJSON(String.Format("Exception: {0} on formula '{1}'", ex.Message, f.Expression)));
                        throw new InvalidPluginExecutionException($"Power Fx error on forumla '{f.Expression}': {ex.Message}");
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
                string jsonStr = JsonConvert.SerializeObject(output); //, Newtonsoft.Json.Formatting.Indented);
                localPluginContext.Trace("JSON Serialised successfuly");

                var response = new pgc_PowerFxEvaluateResponse
                {
                    Results = localPluginContext.PluginExecutionContext.OutputParameters,
                    JSON = jsonStr
                };

                if (request.JSONOnly != true)
                {
                    // Our response will contain a complex JSON schema, so we'll use an expando entity
                    // See: https://powermaverick.dev/2021/11/17/dataverse-custom-api-that-supports-complex-json-schema/#
                    //      https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/web-api-entitytypes#expando
                    //      https://learn.microsoft.com/en-us/dotnet/api/system.dynamic.expandoobject
                    //      https://www.odata.org/getting-started/advanced-tutorial/#openType

                    // First convert our JSON to an expando
                    dynamic expando = JsonConvert.DeserializeObject<ExpandoObject>(jsonStr, new ExpandoObjectConverter());

                    // Then map the expando into a Entity object
                    response.Output = ConvertExpandoToEntity(expando, localPluginContext);
                }
            }
            catch (Exception e)
            {
                localPluginContext.Trace("Got Exception");
                localPluginContext.Trace($"Message: {EscapeJSON(e.Message)}");
                localPluginContext.Trace($"InnerException: {EscapeJSON(e.InnerException.ToString())}");
            }
        }

        private Entity ConvertExpandoToEntity(dynamic expando, ILocalPluginContext localPluginContext)
        {
            var entity = new Entity();
            foreach (KeyValuePair<string, object> kvp in expando)
            {
                localPluginContext.Trace(EscapeJSON(kvp.Key + ": <" + kvp.Value.GetType() + ">" + kvp.Value));

                if (kvp.Value.GetType() == typeof(ExpandoObject))
                {
                    localPluginContext.Trace("New entity needed for " + kvp.Key);
                    entity[kvp.Key] = ConvertExpandoToEntity((ExpandoObject)kvp.Value, localPluginContext);
                }
                else if (kvp.Value.GetType() == typeof(System.Collections.Generic.List<object>))
                {
                    localPluginContext.Trace("Analysing List for " + kvp.Key);
                    var list = (System.Collections.Generic.List<object>)(kvp.Value);

                    // Entity supports StringArray as a property type, so we can return a string array so long
                    // as all members of the list of type string;
                    if (list.All(x => x.GetType() == typeof(string)))
                    {
                        localPluginContext.Trace("List is all strings for " + kvp.Key);
                        entity[kvp.Key] = list.Cast<string>().ToArray();
                    }
                    else
                    {
                        // We can only turn the list into an EntityCollection
                        localPluginContext.Trace("New entity collection needed for " + kvp.Key);
                        var entityCollection = new EntityCollection();

                        foreach (var obj in list)
                        {
                            localPluginContext.Trace("Adding object to entity collection of type " + obj.GetType());
                            if (obj.GetType() == typeof(ExpandoObject))
                            {
                                entityCollection.Entities.Add(ConvertExpandoToEntity((ExpandoObject)obj, localPluginContext));
                            }
                            else
                            {
                                // This should be value types (int, bool, etc), unfortunately we have to create a new entity for it
                                var e = new Entity();
                                e["value"] = obj;
                                entityCollection.Entities.Add(e);
                            }
                        }
                        entity[kvp.Key] = entityCollection;
                    }
                }
                else
                {
                    localPluginContext.Trace("Adding property for " + kvp.Key);
                    entity[kvp.Key] = kvp.Value;
                }
            }
            return entity;
        }

        // Escape JSON so that it can be passed to tracing service which used String.Format()
        private string EscapeJSON(string context)
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

                            localPluginContext.Trace($"expression: {EscapeJSON(expression)}");

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