using System;
using System.Collections;
using System.Collections.Generic;

namespace SC365
{
    // Specifies the config version to use for cmdlets
    public enum MailRouting
    {
        CH,
        PRV
    }

    // Specifies options to be used by cmdlets
    public enum ConfigOption
    {
        Default,
        // DisabledSPFIncoming,
        // DisabledSPFInternal,
        NoAntiSpamWhiteListing
    }

    // Available cloud regions
    public enum GeoRegion
    {
        None,
        CH,
        DE
    }

    public enum ConfigBundle
    {
        None,
        NoTls //Noch nicht aktiv
    }

    // Where should new transport rules be placed, if there are already existing ones
    public enum PlacementPriority
    {
        Top,
        Bottom
    }

    // Not really necessary, but allows fetching and modifying of specific rulesets
    public enum AvailableTransportRuleSettings
    {
        Inbound = 1,
        Outbound = 2,
        Internal = 4,
        EncryptedHeaderCleaning = 8,
        DecryptedHeaderCleaning = 16,
        OutgoingHeaderCleaning = 32,
        All = Inbound | Outbound | Internal | EncryptedHeaderCleaning | DecryptedHeaderCleaning | OutgoingHeaderCleaning
    }

    public enum OperationType
    {
        Create,
        Update
    }

    public static class Bit
    {
        public static ulong Set(ulong no, int pos)
        {return no | ((ulong)1<<pos);}

        public static ulong Clear(ulong no, int pos)
        {return no & (~((ulong)1<<pos));}

        public static ulong Toggle(ulong no, int pos)
        {return no ^ ((ulong)1<<pos);}

        public static ulong Check(ulong no, int pos)
        {return (no>>pos) & 1;}
    }

    public class ConfigBundleSettings
    {
        public ConfigBundleSettings()
        {
            Option = new List<ConfigOption>();
        }

        public ConfigBundleSettings(ConfigBundle id, ConfigVersion version, List<ConfigOption> option)
        {
            Id = id;
            Version = version;
            Option = option;
        }

        public ConfigBundle Id {get; set;}
        public ConfigVersion Version {get; set;}
        public List<ConfigOption> Option {get; set;}
    }

    // We only provide these classes for static typing and to prevent misspells in configuration variables.
    // For convenience they all have a ToHashtable method, in order to use the object with parameter splatting
    public class InboundConnectorSettings
    {
        public InboundConnectorSettings(string name, ConfigVersion version)
        {
            Name = name;
            Version = version;
            Enabled = true;
            EFSkipIPs = new List<string>();
        }

        public ConfigVersion Version {get; private set;}
        public string Name {get; private set;}
        public string Comment {get; set;}
        public string ConnectorSource {get; set;}
        public string ConnectorType {get; set;}
        public string TlsSenderCertificateName {get; set;}
        public bool Skip {get; set;}

        public bool? EFSkipLastIP  {get; set;}
        public bool? RequireTls {get; set;}
        public bool? RestrictDomainsToCertificate {get; set;}
        public bool? RestrictDomainsToIPAddresses {get; set;}
        public bool? CloudServicesMailEnabled {get; set;}
        public bool Enabled {get; set;}

        public List<string> EFUsers {get; set;}
        public List<string> EFSkipIPs {get; set;}
        public List<string> AssociatedAcceptedDomains {get; set;}
        public List<string> SenderDomains {get; set;}

        // This is for splatting
        public Hashtable ToHashtable(OperationType op = OperationType.Create)
        {
            Hashtable ret = new Hashtable();
            ret[(op == OperationType.Create ? "Name" : "Identity")] = Name;
            ret["Enabled"] = Enabled;

            if(!string.IsNullOrEmpty(Comment))
                ret["Comment"] = Comment;
            if(!string.IsNullOrEmpty(ConnectorSource))
                ret["ConnectorSource"] = ConnectorSource;
            if(!string.IsNullOrEmpty(ConnectorType))
                ret["ConnectorType"] = ConnectorType;
            if(!string.IsNullOrEmpty(TlsSenderCertificateName))
                ret["TlsSenderCertificateName"] = TlsSenderCertificateName;

            if(EFSkipLastIP.HasValue)
                ret["EFSkipLastIP"] = EFSkipLastIP.Value;
            if(RequireTls.HasValue)
                ret["RequireTls"] = RequireTls.Value;
            if(RestrictDomainsToCertificate.HasValue)
                ret["RestrictDomainsToCertificate"] = RestrictDomainsToCertificate.Value;
            if(RestrictDomainsToIPAddresses.HasValue)
                ret["RestrictDomainsToIPAddresses"] = RestrictDomainsToIPAddresses.Value;
            if(CloudServicesMailEnabled.HasValue)
                ret["CloudServicesMailEnabled"] = CloudServicesMailEnabled.Value;

            if(EFUsers != null)
                ret["EFUsers"] = EFUsers;
            if(EFSkipIPs != null && EFSkipIPs.Count > 0)
                ret["EFSkipIPs"] = EFSkipIPs;
            if(AssociatedAcceptedDomains != null)
                ret["AssociatedAcceptedDomains"] = AssociatedAcceptedDomains;
            if(SenderDomains != null)
                ret["SenderDomains"] = SenderDomains;

            return ret;
        }
    }

    public class OutboundConnectorSettings
    {
        public OutboundConnectorSettings(string name, ConfigVersion version)
        {
            Name = name;
            Version = version;
            Enabled = true;
        }

        public string Name {get; private set;}
        public ConfigVersion Version {get; private set;}
        public bool Skip {get; set;}

        public string Comment {get; set;}
        public string ConnectorSource {get; set;}
        public string ConnectorType {get; set;}
        public string TlsSettings {get; set;}
        public string TlsDomain {get; set;}

        public bool Enabled {get; set;}
        public bool? IsTransportRuleScoped {get; set;}
        public bool? UseMXRecord {get; set;}
        public bool? CloudServicesMailEnabled {get; set;}

        public List<string> SmartHosts {get; set;}

        // This is for splatting
        public Hashtable ToHashtable(OperationType op = OperationType.Create)
        {
            Hashtable ret = new Hashtable();
            ret[(op == OperationType.Create ? "Name" : "Identity")] = Name;
            ret["Enabled"] = Enabled;

            if(!string.IsNullOrEmpty(Comment))
                ret["Comment"] = Comment;
            if(!string.IsNullOrEmpty(ConnectorSource))
                ret["ConnectorSource"] = ConnectorSource;
            if(!string.IsNullOrEmpty(ConnectorType))
                ret["ConnectorType"] = ConnectorType;
            if(!string.IsNullOrEmpty(TlsSettings))
                ret["TlsSettings"] = TlsSettings;
            if(!string.IsNullOrEmpty(TlsDomain))
                ret["TlsDomain"] = TlsDomain;

            if(IsTransportRuleScoped.HasValue)
                ret["IsTransportRuleScoped"] = IsTransportRuleScoped.Value;
            if(UseMXRecord.HasValue)
                ret["UseMXRecord"] = UseMXRecord.Value;
            if(CloudServicesMailEnabled.HasValue)
                ret["CloudServicesMailEnabled"] = CloudServicesMailEnabled.Value;

            if(SmartHosts != null)
                ret["SmartHosts"] = SmartHosts;

            return ret;
        }
    }

    public class TransportRuleSettings
    {
        public TransportRuleSettings(string name, ConfigVersion version, AvailableTransportRuleSettings type)
        {
            Name = name;
            Version = version;
            Enabled = true;
            Type = type;
        }

        public string Name {get; private set;}
        public ConfigVersion Version {get; private set;}
        public bool Enabled {get; set;}
        public bool Skip {get; set;}
        public AvailableTransportRuleSettings Type {get; private set;}

        public int Priority {get; set;}
        public int SMPriority {get; set;} // used to determine order of SC365 rules

        public int? SetSCL {get; set;}

        public string Comments {get; set;}
        public string FromScope {get; set;}
        public string SentToScope {get; set;}
        public string RouteMessageOutboundConnector {get; set;}
        public string ExceptIfHeaderMatchesMessageHeader {get; set;}
        public string ExceptIfHeaderMatchesPatterns {get; set;}
        public string ExceptIfHeaderContainsMessageHeader {get; set;}
        public string ExceptIfHeaderContainsWords {get; set;}
        public string ExceptIfMessageTypeMatches {get; set;}
        public List<string> ExceptIfRecipientDomainIs {get; set;}
        public List<string> ExceptIfSenderDomainIs {get; set;}
        public string SetAuditSeverity {get; set;}
        public string Mode {get; set;}
        public string SenderAddressLocation {get; set;}
        public string RemoveHeader {get; set;}
        public string HeaderContainsMessageHeader {get; set;}
        public List<string> HeaderContainsWords {get; set;}

        // This is for splatting
        public Hashtable ToHashtable(OperationType op = OperationType.Create)
        {
            Hashtable ret = new Hashtable();
            ret[(op == OperationType.Create ? "Name" : "Identity")] = Name;

            if(op == OperationType.Create)
            {
                ret["Enabled"] = Enabled; // invalid for set
                ret["Priority"] = Priority; // changing priority might be dangerous on set
            }

            if(SetSCL.HasValue)
                ret["SetSCL"] = SetSCL.Value;

            if(!string.IsNullOrEmpty(Comments))
                ret["Comments"] = Comments;
            if(!string.IsNullOrEmpty(FromScope))
                ret["FromScope"] = FromScope;
            if(!string.IsNullOrEmpty(SentToScope))
                ret["SentToScope"] = SentToScope;
            if(op == OperationType.Create &&
                !string.IsNullOrEmpty(RouteMessageOutboundConnector))
                ret["RouteMessageOutboundConnector"] = RouteMessageOutboundConnector;
            if(!string.IsNullOrEmpty(ExceptIfHeaderContainsMessageHeader))
                ret["ExceptIfHeaderContainsMessageHeader"] = ExceptIfHeaderContainsMessageHeader;
            if(!string.IsNullOrEmpty(ExceptIfHeaderContainsWords))
                ret["ExceptIfHeaderContainsWords"] = ExceptIfHeaderContainsWords;
            if(!string.IsNullOrEmpty(ExceptIfHeaderMatchesMessageHeader))
                ret["ExceptIfHeaderMatchesMessageHeader"] = ExceptIfHeaderMatchesMessageHeader;
            if(!string.IsNullOrEmpty(ExceptIfHeaderMatchesPatterns))
                ret["ExceptIfHeaderMatchesPatterns"] = ExceptIfHeaderMatchesPatterns;
            if(!string.IsNullOrEmpty(ExceptIfMessageTypeMatches))
                ret["ExceptIfMessageTypeMatches"] = ExceptIfMessageTypeMatches;
            if(ExceptIfRecipientDomainIs != null)
                ret["ExceptIfRecipientDomainIs"] = ExceptIfRecipientDomainIs;
            if(ExceptIfSenderDomainIs !=null)
                ret["ExceptIfSenderDomainIs"] = ExceptIfSenderDomainIs;
            if(!string.IsNullOrEmpty(SetAuditSeverity))
                ret["SetAuditSeverity"] = SetAuditSeverity;
            if(!string.IsNullOrEmpty(Mode))
                ret["Mode"] = Mode;
            if(!string.IsNullOrEmpty(SenderAddressLocation))
                ret["SenderAddressLocation"] = SenderAddressLocation;
            if(!string.IsNullOrEmpty(RemoveHeader))
                ret["RemoveHeader"] = RemoveHeader;
            if(!string.IsNullOrEmpty(HeaderContainsMessageHeader))
                ret["HeaderContainsMessageHeader"] = HeaderContainsMessageHeader;
            if(HeaderContainsWords != null)
                ret["HeaderContainsWords"] = HeaderContainsWords;

            return ret;
        }
    }

    public class PoliciesAntiSpamSettings
    {
        public PoliciesAntiSpamSettings (string name, GeoRegion georegion)
        {
            Name = name;
            Region = georegion;
        }
        public string Name {get; private set;}
        public GeoRegion Region {get; private set;}
        public bool Skip {get; set;}
        public List<string> WhiteList {get; private set;}

        // This is for splatting
        public Hashtable ToHashtable(OperationType op = OperationType.Create)
        {
            Hashtable ret = new Hashtable();
            ret[(op == OperationType.Create ? "Name" : "Identity")] = Name;
            if(!string.IsNullOrEmpty(WhiteList))
                ret["WhiteList"] = Whitelist;
            return ret;
        }
    }
}
