namespace CustomPipeLineToConsumeExcel.Pipelines
{
    using System;
    using System.Collections.Generic;
    using Microsoft.BizTalk.PipelineOM;
    using Microsoft.BizTalk.Component;
    using Microsoft.BizTalk.Component.Interop;
    
    
    public sealed class ExcelReceive : Microsoft.BizTalk.PipelineOM.ReceivePipeline
    {
        
        private const string _strPipeline = "<?xml version=\"1.0\" encoding=\"utf-16\"?><Document xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instanc"+
"e\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" MajorVersion=\"1\" MinorVersion=\"0\">  <Description /> "+
" <CategoryId>f66b9f5e-43ff-4f5f-ba46-885348ae1b4e</CategoryId>  <FriendlyName>Receive</FriendlyName>"+
"  <Stages>    <Stage>      <PolicyFileStage _locAttrData=\"Name\" _locID=\"1\" Name=\"Decode\" minOccurs=\""+
"0\" maxOccurs=\"-1\" execMethod=\"All\" stageId=\"9d0e4103-4cce-4536-83fa-4a5040674ad6\" />      <Component"+
"s>        <Component>          <Name>PBC.PipelineComponents.XLSX.to.Xml.XlsxtoCSV,CustomPipeLine, Ve"+
"rsion=1.0.0.0, Culture=neutral, PublicKeyToken=a26f3bf5fbcb8075</Name>          <ComponentName>Excel"+
" to xml converter</ComponentName>          <Description>Pipeline component to convert excel file to "+
"xml.</Description>          <Version>1.0.0.0</Version>          <Properties>            <Property Na"+
"me=\"Namespace\" />            <Property Name=\"RootNodeName\" />          </Properties>          <Cache"+
"dDisplayName>Excel to xml converter</CachedDisplayName>          <CachedIsManaged>true</CachedIsMana"+
"ged>        </Component>      </Components>    </Stage>    <Stage>      <PolicyFileStage _locAttrDat"+
"a=\"Name\" _locID=\"2\" Name=\"Disassemble\" minOccurs=\"0\" maxOccurs=\"-1\" execMethod=\"FirstMatch\" stageId="+
"\"9d0e4105-4cce-4536-83fa-4a5040674ad6\" />      <Components />    </Stage>    <Stage>      <PolicyFil"+
"eStage _locAttrData=\"Name\" _locID=\"3\" Name=\"Validate\" minOccurs=\"0\" maxOccurs=\"-1\" execMethod=\"All\" "+
"stageId=\"9d0e410d-4cce-4536-83fa-4a5040674ad6\" />      <Components />    </Stage>    <Stage>      <P"+
"olicyFileStage _locAttrData=\"Name\" _locID=\"4\" Name=\"ResolveParty\" minOccurs=\"0\" maxOccurs=\"-1\" execM"+
"ethod=\"All\" stageId=\"9d0e410e-4cce-4536-83fa-4a5040674ad6\" />      <Components />    </Stage>  </Sta"+
"ges></Document>";
        
        private const string _versionDependentGuid = "154e0ea9-aaa0-4c90-bf7b-5b355cdb3c1d";
        
        public ExcelReceive()
        {
            Microsoft.BizTalk.PipelineOM.Stage stage = this.AddStage(new System.Guid("9d0e4103-4cce-4536-83fa-4a5040674ad6"), Microsoft.BizTalk.PipelineOM.ExecutionMode.all);
            IBaseComponent comp0 = Microsoft.BizTalk.PipelineOM.PipelineManager.CreateComponent("PBC.PipelineComponents.XLSX.to.Xml.XlsxtoCSV,CustomPipeLine, Version=1.0.0.0, Culture=neutral, PublicKeyToken=a26f3bf5fbcb8075");;
            if (comp0 is IPersistPropertyBag)
            {
                string comp0XmlProperties = "<?xml version=\"1.0\" encoding=\"utf-16\"?><PropertyBag xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-inst"+
"ance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">  <Properties>    <Property Name=\"Namespace\" />  "+
"  <Property Name=\"RootNodeName\" />  </Properties></PropertyBag>";
                PropertyBag pb = PropertyBag.DeserializeFromXml(comp0XmlProperties);;
                ((IPersistPropertyBag)(comp0)).Load(pb, 0);
            }
            this.AddComponent(stage, comp0);
        }
        
        public override string XmlContent
        {
            get
            {
                return _strPipeline;
            }
        }
        
        public override System.Guid VersionDependentGuid
        {
            get
            {
                return new System.Guid(_versionDependentGuid);
            }
        }
    }
}
