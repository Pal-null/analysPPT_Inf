using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

using System.Data;
using System.Data.OracleClient;

using System.Xml;

namespace PPT_0309
{
    public interface PPT_0309_1Ifce
    {
        void doMain(int _paperId, int _stuId, string _docName, string _xmlPath);
    };
    public class getPptResult : PPT_0309_1Ifce
    {
        public void doMain(int _paperId, int _stuId, string _docName, string _xmlPath)
        {
            new Analys().analys(_paperId, _stuId, _docName, _xmlPath);
            Console.WriteLine("PPT Analys SUCCESS----------CSharp");
        }
    }

    public class Analys
    {
        #region 全局变量
        private int paperID;
        private int stuID;
        private string docName;
        private string savePath;
        private string xmlPath;

        private int rootID = 0;
        private String fileNodeName;
        private String xmlFileName;
        private int attrID = 0;
        private int imageIndex = 0;
        //文件名编号
        private int c_slides = 0;
        private int c_notesSlides = 0;
        private int c_slideMasters = 0;
        private int c_notesMasters = 0;
        private int c_theme = 0;
        private int c_slideLayouts = 0;
        private int c_presentationPr = 0;
        private int c_tblStyleLst = 0;
        private int c_viewPr = 0;
        private int c_handoutMaster = 0;

        private XmlDocument docNode;
        private XmlElement RootNode;
        private XmlDocument docAttr;
        private XmlElement RootAttr;

        private OracleConnection oracleConn;
        #endregion

        #region 解析接口
        public void analys(int _paperId, int _stuId, string _docName, string _xmlPath)
        {
            paperID = _paperId;
            stuID = _stuId;
            docName = _docName;
            xmlPath = _xmlPath;
            savePath = xmlPath + "images/";

            docNode = null;
            RootNode = null;
            docAttr = null;
            RootAttr = null;

            if (!System.IO.Directory.Exists(savePath))
            {
                System.IO.DirectoryInfo directoryInfo = new System.IO.DirectoryInfo(savePath);
                directoryInfo.Create();
            }

            docNode = new XmlDocument();
            RootNode = docNode.CreateElement("Root");
            docAttr = new XmlDocument();
            RootAttr = docAttr.CreateElement("Root");

            docNode.AppendChild(RootNode);
            docAttr.AppendChild(RootAttr);

            XmlElement totalScore = docAttr.CreateElement("totalScore");
            totalScore.InnerText = "0";
            RootAttr.AppendChild(totalScore);

            oracleConn = getOracleConn("localhost", "1521", "orcl", "root", "root");
            try
            {
                oracleConn.Open();
            }
            catch (Exception ex)
            {
                Console.WriteLine("数据库连接打开失败：" + ex.Message);
            }

            //开始解析
            GetResult(docName);

            oracleConn.Close();

            docNode.Save(xmlPath + paperID.ToString() + "-" + stuID.ToString() + "-" + "node.xml");
            docAttr.Save(xmlPath + paperID.ToString() + "-" + stuID.ToString() + "-" + "attr.xml");
        }
        #endregion

        #region 递归解析
        public void GetResult(string docName)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(docName, false))
            {
                IEnumerable<IdPartPair> Presentation_parts = ppt.PresentationPart.Parts;
                int slideID;
                if (ppt.PresentationPart.SlideParts.Count() > 0)
                {
                    slideID = ++rootID;
                    writeNodeToXML(slideID, 0, "幻灯片", "", "slide/", "false");
                }
                slideID = rootID;
                int slideMasterID;
                if (ppt.PresentationPart.SlideMasterParts.Count() > 0)
                {
                    slideMasterID = ++rootID;
                    writeNodeToXML(slideMasterID, 0, "幻灯片母版", "", "slideMaster/", "false");
                }
                slideMasterID = rootID;
                int notesMasterID;
                if (ppt.PresentationPart.NotesMasterPart != null)
                {
                    notesMasterID = ++rootID;
                    writeNodeToXML(notesMasterID, 0, "备注母版", "", "notesMaster/", "false");
                }
                notesMasterID = rootID;
                int themeID;
                if (ppt.PresentationPart.ThemePart != null)
                {
                    themeID = ++rootID;
                    writeNodeToXML(themeID, 0, "主题", "", "theme/", "false");
                }
                themeID = rootID;
                int presentationID = ++rootID;
                writeNodeToXML(presentationID, 0, "演示文稿概览", "", "presentation/", "false");
                int presentationPrID = ++rootID;
                writeNodeToXML(presentationPrID, 0, "演示文稿属性", "", "presentationProperties/", "false");
                int tblStyleLstID = ++rootID;
                writeNodeToXML(tblStyleLstID, 0, "表格样式列表", "", "tableStyleList/", "false");
                int viewPrID = ++rootID;
                writeNodeToXML(viewPrID, 0, "视图属性", "", "viewProperties/", "false");
                int handoutMasterID;
                if (ppt.PresentationPart.HandoutMasterPart != null && ppt.PresentationPart.HandoutMasterPart.Parts.Count() > 0)
                {
                    handoutMasterID = ++rootID;
                    writeNodeToXML(handoutMasterID, 0, "讲义母版", "", "handoutMaster/", "false");
                }
                handoutMasterID = rootID;
                int extendedFilePropertiesID = ++rootID;
                writeNodeToXML(extendedFilePropertiesID, 0, "扩展文件属性", "", "extendedFileProperties/", "false");

                #region Presentation部分
                for (int i = 1; i <= Presentation_parts.ToArray().Length; i++)
                {
                    OpenXmlPart part = ppt.PresentationPart.GetPartById("rId" + i);
                    fileNodeName = part.RootElement.LocalName;
                    #region 幻灯片部分
                    if (fileNodeName == "sld")
                    {
                        c_slides++;
                        xmlFileName = "幻灯片" + c_slides;
                        SlidePart p = (SlidePart)part;
                        int CurrentRootID = rootID;
                        if (p.RootElement != null)
                        {
                            getAttribute(p.RootElement, 2, c_slides, slideID, "slide/" + xmlFileName + "/", c_slides, p);
                        }
                        if (p.Parts.Count() > 0)
                        {
                            int wi;
                            writeNodeToXML(++rootID, CurrentRootID + 1, "关联", "", "slide/" + xmlFileName + "/rId", "true");
                            for (wi = 1; wi <= p.Parts.Count(); wi++)
                            {
                                writeAttrToXML(++attrID, 0, "rId" + wi, p.GetPartById("rId" + wi).Uri.ToString(), "slide/" + xmlFileName + "/rId" + wi, "0", "0", "null");
                            }
                        }
                        if (p.SlideLayoutPart.RootElement != null)
                        {
                            writeNodeToXML(++rootID, ++CurrentRootID, "幻灯片版式", "", "slide/" + xmlFileName + "/slideLayout/", "false");
                            getAttribute(p.SlideLayoutPart.RootElement, 3, 7, rootID, "slide/" + xmlFileName + "/slideLayout/", 1, null);
                        }
                        if (p.NotesSlidePart != null && p.NotesSlidePart.RootElement != null)
                        {
                            writeNodeToXML(++rootID, CurrentRootID, "备注幻灯片", "", "slide/" + xmlFileName + "/notesSlide/", "false");
                            getAttribute(p.NotesSlidePart.RootElement, 3, 8, rootID, "slide/" + xmlFileName + "/notesSlide/", 1, null);
                        }
                        if (p.SlideCommentsPart != null && p.SlideCommentsPart.RootElement != null)
                        {
                            writeNodeToXML(++rootID, CurrentRootID, "幻灯片批注", "", "slide/" + xmlFileName + "/slideComments/", "false");
                            getAttribute(p.SlideCommentsPart.RootElement, 3, 9, rootID, "slide/" + xmlFileName + "/slideComments/", 1, null);
                        }
                        if (p.ChartParts.Count() > 0)
                        {
                            int chartPartCount = 0;
                            foreach (ChartPart chartPart in p.ChartParts)
                            {
                                //图表主体
                                if (chartPart.RootElement.LocalName == "chartSpace")
                                {
                                    writeNodeToXML(++rootID, CurrentRootID, "图表空间", "", "slide/" + xmlFileName + "/chartSpace/", "false");
                                    getAttribute(chartPart.RootElement, 3, 10, rootID, "slide/" + xmlFileName + "/chartSpace/", ++chartPartCount, null);
                                }
                                //图表样式
                                Hashtable hashTable = new Hashtable();
                                foreach (ChartStylePart stylePart in chartPart.ChartStyleParts)
                                {
                                    if (hashTable.Contains(stylePart.RootElement.ToString()))
                                    {
                                        int ii = (int)hashTable[stylePart.RootElement.ToString()] + 1;
                                        hashTable.Remove(stylePart.RootElement.ToString());
                                        hashTable.Add(stylePart.RootElement.ToString(), ii);
                                    }
                                    else
                                    {
                                        hashTable.Add(stylePart.RootElement.ToString(), 1);
                                    }
                                    writeNodeToXML(++rootID, CurrentRootID, "图标样式", "", "slide/" + xmlFileName + "/chartStyle/", "false");
                                    getAttribute(stylePart.RootElement, 4, (int)hashTable[stylePart.RootElement.ToString()], rootID, "slide/" + xmlFileName + "/chartStyle/", (int)hashTable[stylePart.RootElement.ToString()], null);
                                }
                                //图表
                                hashTable.Clear();
                                foreach (ChartColorStylePart colorStylePart in chartPart.ChartColorStyleParts)
                                {
                                    if (hashTable.Contains(colorStylePart.RootElement.ToString()))
                                    {
                                        int ii = (int)hashTable[colorStylePart.RootElement.ToString()] + 1;
                                        hashTable.Remove(colorStylePart.RootElement.ToString());
                                        hashTable.Add(colorStylePart.RootElement.ToString(), ii);
                                    }
                                    else
                                    {
                                        hashTable.Add(colorStylePart.RootElement.ToString(), 1);

                                    }
                                    writeNodeToXML(++rootID, CurrentRootID, "图标颜色风格", "", "slide/" + xmlFileName + "/chartColorStyle/", "false");
                                    getAttribute(colorStylePart.RootElement, 4, (int)hashTable[colorStylePart.RootElement.ToString()], rootID, "slide/" + xmlFileName + "/chartColorStyle/", (int)hashTable[colorStylePart.RootElement.ToString()], null);
                                }
                                //图表图片
                                if (chartPart.ChartDrawingPart != null && chartPart.ChartDrawingPart.RootElement != null)
                                {
                                    writeNodeToXML(CurrentRootID, ++rootID, get_typeName(chartPart.ChartDrawingPart.RootElement.GetType().ToString()), "", "slide/" + xmlFileName + "/chartSpace/", "true");
                                }
                            }
                        }
                        //示意图
                        if (p.DiagramColorsParts.Count() > 0)
                        {
                            Hashtable hashTable = new Hashtable();
                            foreach (DiagramColorsPart diagramColorsPart in p.DiagramColorsParts)
                            {
                                if (hashTable.Contains(diagramColorsPart.RootElement.ToString()))
                                {
                                    int ii = (int)hashTable[diagramColorsPart.RootElement.ToString()] + 1;
                                    hashTable.Remove(diagramColorsPart.RootElement.ToString());
                                    hashTable.Add(diagramColorsPart.RootElement.ToString(), ii);
                                }
                                else
                                {
                                    hashTable.Add(diagramColorsPart.RootElement.ToString(), 1);
                                }
                                writeNodeToXML(++rootID, CurrentRootID, "示意图颜色映射", "", "slide/" + xmlFileName + "/diagramColors/", "false");
                                getAttribute(diagramColorsPart.RootElement, 4, (int)hashTable[diagramColorsPart.RootElement.ToString()], rootID, "slide/" + xmlFileName + "/diagramColors/", (int)hashTable[diagramColorsPart.RootElement.ToString()], null);
                            }
                        }
                        if (p.DiagramDataParts.Count() > 0)
                        {
                            Hashtable hashTable = new Hashtable();
                            foreach (DiagramDataPart diagramDataPart in p.DiagramDataParts)
                            {
                                if (hashTable.Contains(diagramDataPart.RootElement.ToString()))
                                {
                                    int ii = (int)hashTable[diagramDataPart.RootElement.ToString()] + 1;
                                    hashTable.Remove(diagramDataPart.RootElement.ToString());
                                    hashTable.Add(diagramDataPart.RootElement.ToString(), ii);
                                }
                                else
                                {
                                    hashTable.Add(diagramDataPart.RootElement.ToString(), 1);
                                }
                                writeNodeToXML(++rootID, CurrentRootID, "示意图数据", "", "slide/" + xmlFileName + "/diagramData/", "false");
                                getAttribute(diagramDataPart.RootElement, 4, (int)hashTable[diagramDataPart.RootElement.ToString()], rootID, "slide/" + xmlFileName + "/diagramData/", (int)hashTable[diagramDataPart.RootElement.ToString()], null);
                            }
                        }
                        if (p.DiagramStyleParts.Count() > 0)
                        {
                            Hashtable hashTable = new Hashtable();
                            foreach (DiagramStylePart diagramStylePart in p.DiagramStyleParts)
                            {
                                if (hashTable.Contains(diagramStylePart.RootElement.ToString()))
                                {
                                    int ii = (int)hashTable[diagramStylePart.RootElement.ToString()] + 1;
                                    hashTable.Remove(diagramStylePart.RootElement.ToString());
                                    hashTable.Add(diagramStylePart.RootElement.ToString(), ii);
                                }
                                else
                                {
                                    hashTable.Add(diagramStylePart.RootElement.ToString(), 1);
                                }
                                writeNodeToXML(++rootID, CurrentRootID, "示意图样式", "", "slide/" + xmlFileName + "/diagramStyle/", "false");
                                getAttribute(diagramStylePart.RootElement, 4, (int)hashTable[diagramStylePart.RootElement.ToString()], rootID, "slide/" + xmlFileName + "/diagramStyle/", (int)hashTable[diagramStylePart.RootElement.ToString()], null);
                            }
                        }
                        if (p.DiagramPersistLayoutParts.Count() > 0)
                        {
                            Hashtable hashTable = new Hashtable();
                            foreach (DiagramPersistLayoutPart diagramPersistLayoutPart in p.DiagramPersistLayoutParts)
                            {
                                if (hashTable.Contains(diagramPersistLayoutPart.RootElement.ToString()))
                                {
                                    int ii = (int)hashTable[diagramPersistLayoutPart.RootElement.ToString()] + 1;
                                    hashTable.Remove(diagramPersistLayoutPart.RootElement.ToString());
                                    hashTable.Add(diagramPersistLayoutPart.RootElement.ToString(), ii);
                                }
                                else
                                {
                                    hashTable.Add(diagramPersistLayoutPart.RootElement.ToString(), 1);
                                }
                                writeNodeToXML(++rootID, CurrentRootID, "示意图保存样式", "", "slide/" + xmlFileName + "/diagramPersistLayout/", "false");
                                getAttribute(diagramPersistLayoutPart.RootElement, 4, (int)hashTable[diagramPersistLayoutPart.RootElement.ToString()], rootID, "slide/" + xmlFileName + "/diagramPersistLayout/", (int)hashTable[diagramPersistLayoutPart.RootElement.ToString()], null);
                            }
                        }
                        if (p.DiagramLayoutDefinitionParts.Count() > 0)
                        {
                            Hashtable hashTable = new Hashtable();
                            foreach (DiagramLayoutDefinitionPart diagramLayoutDefinitionPart in p.DiagramLayoutDefinitionParts)
                            {
                                if (hashTable.Contains(diagramLayoutDefinitionPart.RootElement.ToString()))
                                {
                                    int ii = (int)hashTable[diagramLayoutDefinitionPart.RootElement.ToString()] + 1;
                                    hashTable.Remove(diagramLayoutDefinitionPart.RootElement.ToString());
                                    hashTable.Add(diagramLayoutDefinitionPart.RootElement.ToString(), ii);
                                }
                                else
                                {
                                    hashTable.Add(diagramLayoutDefinitionPart.RootElement.ToString(), 1);
                                }
                                writeNodeToXML(++rootID, CurrentRootID, "示意图样式定义", "", "slide/" + xmlFileName + "/diagramLayoutDefinition/", "false");
                                getAttribute(diagramLayoutDefinitionPart.RootElement, 4, (int)hashTable[diagramLayoutDefinitionPart.RootElement.ToString()], rootID, "slide/" + xmlFileName + "/diagramLayoutDefinition/", (int)hashTable[diagramLayoutDefinitionPart.RootElement.ToString()], null);
                            }
                        }
                    }
                    #endregion

                    else if (fileNodeName == "sldMaster")
                    {
                        c_slideMasters++;
                        xmlFileName = "幻灯片母版" + c_slideMasters;
                        SlideMasterPart p = (SlideMasterPart)part;
                        if (p.RootElement != null)
                        {
                            int CurrentRootID = rootID + 1;
                            getAttribute(p.RootElement, 2, c_slideMasters, slideMasterID, "slideMaster/" + xmlFileName + "/", c_slideMasters, null);
                            if (p.Parts.Count() > 0)
                            {
                                int wi;
                                writeNodeToXML(++rootID, CurrentRootID, "关联", "", "slideMaster/" + xmlFileName + "/" + "rId", "true");
                                for (wi = 1; wi <= p.Parts.Count(); wi++)
                                {
                                    writeAttrToXML(++attrID, 0, "rId" + wi, p.GetPartById("rId" + wi).Uri.ToString(), "slideMaster/" + xmlFileName + "/" + "rId" + wi, "0", "0", "null");
                                }
                            }
                        }
                    }
                    else if (fileNodeName == "notesMaster")
                    {
                        c_notesMasters++;
                        xmlFileName = "notesMaster" + c_notesMasters;
                        NotesMasterPart p = (NotesMasterPart)part;
                        if (p.RootElement != null)
                        {
                            getAttribute(p.RootElement, 2, c_notesMasters, notesMasterID, "notesMaster/" + xmlFileName + "/", c_notesMasters, null);
                        }
                    }
                    else if (fileNodeName == "theme")
                    {
                        c_theme++;
                        xmlFileName = "theme" + c_theme;
                        ThemePart p = (ThemePart)part;
                        if (p.RootElement != null)
                        {
                            getAttribute(p.RootElement, 2, c_theme, themeID, "theme/" + xmlFileName + "/", c_theme, null);
                        }
                    }
                    else if (fileNodeName == "presentationPr")
                    {
                        c_presentationPr++;
                        xmlFileName = "presentationPr" + c_presentationPr;
                        PresentationPropertiesPart p = (PresentationPropertiesPart)part;
                        if (p.RootElement != null)
                        {
                            getAttribute(p.RootElement, 2, c_presentationPr, presentationPrID, "presentationProperties/" + xmlFileName + "/", c_presentationPr, null);
                        }
                    }
                    else if (fileNodeName == "tblStyleLst")
                    {
                        c_tblStyleLst++;
                        TableStylesPart p = (TableStylesPart)part;
                        if (p.RootElement != null)
                        {
                            getAttribute(p.RootElement, 2, c_tblStyleLst, tblStyleLstID, "tableStyleList/", c_tblStyleLst, null);
                        }
                    }
                    else if (fileNodeName == "viewPr")
                    {
                        c_viewPr++;
                        ViewPropertiesPart p = (ViewPropertiesPart)part;
                        if (p.RootElement != null)
                        {
                            getAttribute(p.RootElement, 2, c_viewPr, viewPrID, "viewProperties/", c_viewPr, null);
                        }
                    }
                    else if (fileNodeName == "handoutMaster")
                    {
                        c_handoutMaster++;
                        HandoutMasterPart p = (HandoutMasterPart)part;
                        if (p.RootElement != null)
                        {
                            getAttribute(p.RootElement, 2, c_handoutMaster, handoutMasterID, "handoutMaster/", c_handoutMaster, null);
                        }
                    }
                }
                #endregion

                #region Presentation.xml
                int curID = rootID + 1;
                getAttribute(ppt.PresentationPart.Presentation.PresentationPart.RootElement, 2, 1, presentationID, "presentation/", 1, null);
                if (ppt.PresentationPart.Presentation.PresentationPart.Parts.Count() > 0)
                {
                    writeNodeToXML(++rootID, curID, "关联", "", "presentation/rId", "true");
                    int wi;
                    for (wi = 1; wi <= ppt.PresentationPart.Presentation.PresentationPart.Parts.Count(); wi++)
                    {
                        writeAttrToXML(++attrID, 0, "rId" + wi, ppt.PresentationPart.Presentation.PresentationPart.GetPartById("rId" + wi).Uri.ToString(), "presentation/rId" + wi, "0", "0", "null");
                    }
                }
                #endregion

                //#region CoreFileProperties部分
                ////getAttribute(ppt.CoreFilePropertiesPart.RootElement, 0, 1, 1, "coreProperties19");
                //#endregion

                #region ExtendedFileProperties部分
                getAttribute(ppt.ExtendedFilePropertiesPart.Properties, 2, 1, extendedFilePropertiesID, "extendedFileProperties/", 1, null);
                #endregion

                //#region Thumbnail部分
                ////addRow_WtreeAttrs("首页预览图", ppt.ThumbnailPart.Uri.ToString(), "thumbnai21", "0", "0", 0, 1, 1);
                //#endregion
            }
        }
        #endregion

        #region 获取所有属性
        public void getAttribute(OpenXmlElement element, int depth, int serial, int fatherID, String prefix, int nodeCount, SlidePart thisSlide)
        {
            depth++;
            rootID++;
            prefix += element.LocalName + nodeCount + "/";
            int thisID = rootID;
            bool hasChildren = element.HasChildren;
            bool hasAttributes = element.HasAttributes;

            //如果此节点有子节点但没有属性
            if (hasChildren && !hasAttributes)
            {
                writeNodeToXML(thisID, fatherID, get_typeName(element.GetType().ToString()) + nodeCount, element.InnerText, prefix, "false");
                //判断是否是图片
                if (element.LocalName == "pic")
                {
                    ImagePart imagePart = (ImagePart)thisSlide.GetPartById(element.GetFirstChild<BlipFill>().Blip.Embed);
                    System.Drawing.Image img = System.Drawing.Image.FromStream(imagePart.GetStream());
                    imageIndex++;
                    String fileName = paperID + stuID + rootID + "image" + imageIndex + ".gif";
                    img.Save(savePath + fileName, System.Drawing.Imaging.ImageFormat.Gif);
                    writeAttrToXML(++attrID, 0, "资源文件", fileName, prefix, "0", "0", "null");
                }
                else if (element.LocalName == "transition")
                {
                    writeAttrToXML(++attrID, 0, "切换效果", element.LocalName, prefix + element.FirstChild.LocalName + "1/", "0", "0", "null");
                }
                int serial_child = 1;
                Hashtable hashTable = new Hashtable();
                foreach (OpenXmlElement e in element.ChildElements)
                {
                    if (hashTable.Contains(e.LocalName))
                    {
                        int i = (int)hashTable[e.LocalName] + 1;
                        hashTable.Remove(e.LocalName);
                        hashTable.Add(e.LocalName, i);
                    }
                    else
                    {
                        hashTable.Add(e.LocalName, 1);
                    }
                    getAttribute(e, depth, serial_child, thisID, prefix, (int)hashTable[e.LocalName], thisSlide);
                    serial_child++;
                }
                return;
            }
            //如果此节点既没有属性也没有子节点
            else if (!hasAttributes && !hasChildren)
            {
                writeNodeToXML(thisID, fatherID, get_typeName(element.GetType().ToString()) + nodeCount, element.InnerText, prefix, "true");
                writeAttrToXML(+attrID, 0, get_typeName(element.GetType().ToString()), element.InnerText, prefix, "0", "0", "null");
                return;
            }
            //如果此节点有属性且有子节点
            else if (hasAttributes && hasChildren)
            {
                if (element.LocalName == "transition")
                {
                    writeAttrToXML(++attrID, 0, "切换效果", element.LocalName, prefix + element.FirstChild.LocalName + "1/", "0", "0", "null");
                }
                writeNodeToXML(thisID, fatherID, get_typeName(element.GetType().ToString()) + nodeCount, element.InnerText, prefix, "false");
                foreach (OpenXmlAttribute attr in element.GetAttributes())
                {
                    writeAttrToXML(++attrID, 0, get_attrChinese(element.GetType().ToString(), attr.LocalName), attr.Value, prefix, "0", "0", "null");
                }
                int serial_child = 1;
                Hashtable hashTable = new Hashtable();
                foreach (OpenXmlElement e in element.ChildElements)
                {
                    if (hashTable.Contains(e.LocalName))
                    {
                        int i = (int)hashTable[e.LocalName] + 1;
                        hashTable.Remove(e.LocalName);
                        hashTable.Add(e.LocalName, i);
                    }
                    else
                    {
                        hashTable.Add(e.LocalName, 1);
                    }
                    getAttribute(e, depth, serial_child, thisID, prefix, (int)hashTable[e.LocalName], thisSlide);
                    serial_child++;
                }
                return;
            }
            //如果有属性但无子节点
            else if (hasAttributes && !hasChildren)
            {
                writeNodeToXML(thisID, fatherID, get_typeName(element.GetType().ToString()) + nodeCount, element.InnerText, prefix, "true");
                foreach (OpenXmlAttribute attr in element.GetAttributes())
                {
                    writeAttrToXML(++attrID, 0, get_attrChinese(element.GetType().ToString(), attr.LocalName), attr.Value, prefix, "0", "0", "null");
                }
                return;
            }
        }
        #endregion

        #region 数据库操作部分
        public OracleConnection getOracleConn(String Host, String Port, String serviceName, String UserID, String Password)
        {
            OracleConnectionStringBuilder OcnnStrB = new OracleConnectionStringBuilder();
            OcnnStrB.DataSource = "(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" + Host + ") (PORT=" + Port + ")))(CONNECT_DATA=(SERVICE_NAME=" + serviceName + ")))";
            OcnnStrB.UserID = UserID;
            OcnnStrB.Password = Password;
            OracleConnection myCnn = new OracleConnection(OcnnStrB.ConnectionString);
            return myCnn;
        }
        #endregion

        #region 获取中文翻译
        String get_typeName(String elementType)
        {
            String[] arry = elementType.Split('.');
            String className = arry[arry.Length - 1];
            String nameSpace = arry[0];
            int i;
            for (i = 1; i < arry.Length - 1; i++)
            {
                nameSpace += "." + arry[i];
            }
            OracleCommand com = oracleConn.CreateCommand();
            com.CommandText = "select TRANSLATION from TRANSLATE_NODE where NAMESPACE=\'" + nameSpace +
                "\' and CLASS_NAME=\'" + className + "\'";
            OracleDataReader odr = com.ExecuteReader();
            if (odr.Read())
            {
                String odrString = odr.GetString(0).ToString();
                odr.Close();
                return odrString;
            }
            else
            {
                return className;
            }
        }

        String get_attrChinese(String elementType, String localName)
        {
            String[] arry = elementType.Split('.');
            String className = arry[arry.Length - 1];
            String nameSpace = arry[0];
            int i;
            for (i = 1; i < arry.Length - 1; i++)
            {
                nameSpace += "." + arry[i];
            }
            OracleCommand com = oracleConn.CreateCommand();
            com.CommandText = "select TRANSLATION from TRANSLATE_ATTR where NAMESPACE=\'" + nameSpace +
                "\' and CLASS_NAME=\'" + className + "\' and ATTR_NAME= '" + localName + "\'";
            OracleDataReader odr = com.ExecuteReader();
            if (odr.Read())
            {
                String odrString = odr.GetString(0).ToString();
                odr.Close();
                return odrString;
            }
            else
            {
                return localName;
            }
        }
        #endregion

        #region 存入XML文件
        public void writeNodeToXML(int ID, int fatherID, String elementName, String Content, String Prefix, String leaf)
        {
            XmlElement element = docNode.CreateElement("record");
            element.SetAttribute("ID", ID.ToString());
            element.SetAttribute("fid", fatherID.ToString());            
            element.SetAttribute("node", elementName);
            element.SetAttribute("content", Content);
            element.SetAttribute("prefix", Prefix);
            element.SetAttribute("leaf", leaf);
            element.SetAttribute("paper", paperID.ToString());
            element.SetAttribute("userid", stuID.ToString());
            RootNode.AppendChild(element);
        }

        public void writeAttrToXML(int ID, int fatherID, String attrName, String value, String Prefix,
            String score, String status, String checkType)
        {
            XmlElement element = docAttr.CreateElement("record");
            element.SetAttribute("ID", ID.ToString());
            element.SetAttribute("fid", fatherID.ToString());
            element.SetAttribute("prefix", Prefix);
            element.SetAttribute("attr", attrName);
            element.SetAttribute("value", value);
            element.SetAttribute("score", score);
            element.SetAttribute("status", status);
            element.SetAttribute("checkType", checkType);
            element.SetAttribute("paper", paperID.ToString());
            element.SetAttribute("userid", stuID.ToString());
            RootAttr.AppendChild(element);
        }
        #endregion
    }
}
