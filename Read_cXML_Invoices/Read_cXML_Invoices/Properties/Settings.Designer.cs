//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Read_cXML_Invoices.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "17.2.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.WebServiceUrl)]
        [global::System.Configuration.DefaultSettingValueAttribute("http://172.16.25.80:7047/DynamicsNAV90/WS/Government%20Scientific%20Source/Codeun" +
            "it/AutoPostDocument")]
        public string Read_cXML_Invoices_PrdAutoPostDoc_AutoPostDocument {
            get {
                return ((string)(this["Read_cXML_Invoices_PrdAutoPostDoc_AutoPostDocument"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.WebServiceUrl)]
        [global::System.Configuration.DefaultSettingValueAttribute("http://172.16.25.80:7047/DynamicsNAV90/WS/Government%20Scientific%20Source/Page/P" +
            "urchaseOrder")]
        public string Read_cXML_Invoices_PrdPurchaseOrder_PurchaseOrder_Service {
            get {
                return ((string)(this["Read_cXML_Invoices_PrdPurchaseOrder_PurchaseOrder_Service"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.WebServiceUrl)]
        [global::System.Configuration.DefaultSettingValueAttribute("http://172.16.25.121:7047/DynamicsNAV90/WS/Government%20Scientific%20Source/Codeu" +
            "nit/AutoPostDocument")]
        public string Read_cXML_Invoices_DevZAutoPostDoc_AutoPostDocument {
            get {
                return ((string)(this["Read_cXML_Invoices_DevZAutoPostDoc_AutoPostDocument"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.WebServiceUrl)]
        [global::System.Configuration.DefaultSettingValueAttribute("http://172.16.25.121:7047/DynamicsNAV90/WS/Government%20Scientific%20Source/Page/" +
            "PurchaseOrder")]
        public string Read_cXML_Invoices_DevZPurchaseOrder_PurchaseOrder_Service {
            get {
                return ((string)(this["Read_cXML_Invoices_DevZPurchaseOrder_PurchaseOrder_Service"]));
            }
        }
    }
}
