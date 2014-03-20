﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

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
        private static int paperID;
        private static int stuID;
        private static string docName;
        private static string savePath;
        private static string xmlPath;

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

        private static XmlDocument docNode = null;
        private static XmlElement RootNode = null;
        private static XmlDocument docAttr = null;
        private static XmlElement RootAttr = null;
        #endregion

        #region 解析接口
        public void analys(int _paperId, int _stuId, string _docName, string _xmlPath)
        {
            paperID = _paperId;
            stuID = _stuId;
            docName = _docName;
            xmlPath = _xmlPath;
            savePath = xmlPath + "images/";
            Console.WriteLine(docName);
            Console.WriteLine(xmlPath);
            Console.WriteLine(savePath);
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

            //开始解析
            GetResult(docName);

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
                    //addRow_Wtree("幻灯片", "slides", 0, slideID, "slide/", "0", "", 1, 1, "0");
                }
                slideID = rootID;
                int slideMasterID;
                if (ppt.PresentationPart.SlideMasterParts.Count() > 0)
                {
                    slideMasterID = ++rootID;
                    writeNodeToXML(slideMasterID, 0, "幻灯片母版", "", "slideMaster/", "false");
                    //addRow_Wtree("幻灯片母版", "sldMasters", 0, slideMasterID, "slideMaster/", "0", "", 1, 2, "0");
                }
                slideMasterID = rootID;
                int notesMasterID;
                if (ppt.PresentationPart.NotesMasterPart != null)
                {
                    notesMasterID = ++rootID;
                    writeNodeToXML(notesMasterID, 0, "notesMaster", "", "notesMaster/", "false");
                    //addRow_Wtree("备注母版", "notesMaster", 0, notesMasterID, "notesMaster/", "0", "", 1, 3, "0");
                }
                notesMasterID = rootID;
                int themeID;
                if (ppt.PresentationPart.ThemePart != null)
                {
                    themeID = ++rootID;
                    writeNodeToXML(themeID, 0, "主题", "", "theme/", "false");
                    //addRow_Wtree("主题", "theme", 0, themeID, "theme/", "0", "", 1, 4, "0");
                }
                themeID = rootID;
                int presentationID = ++rootID;
                writeNodeToXML(presentationID, 0, "演示文稿概览", "", "presentation/", "false");
                //addRow_Wtree("演示文稿概览", "presentation", 0, presentationID, "presentation/", "0", "", 1, 5, "0");
                int presentationPrID = ++rootID;
                writeNodeToXML(presentationPrID, 0, "演示文稿属性", "", "presentationProperties/", "false");
                //addRow_Wtree("演示文稿属性", "presentationPr", 0, presentationPrID, "presentationProperties/", "0", "", 1, 6, "0");
                int tblStyleLstID = ++rootID;
                writeNodeToXML(tblStyleLstID, 0, "表格样式列表", "", "tableStyleList/", "false");
                //addRow_Wtree("表格样式列表", "tblStyleLst", 0, tblStyleLstID, "tableStyleList/", "0", "", 1, 7, "0");
                int viewPrID = ++rootID;
                writeNodeToXML(viewPrID, 0, "视图属性", "", "viewProperties/", "false");
                //addRow_Wtree("视图属性", "viewPr", 0, viewPrID, "viewProperties/", "0", "", 1, 8, "0");
                int handoutMasterID;
                if (ppt.PresentationPart.HandoutMasterPart != null && ppt.PresentationPart.HandoutMasterPart.Parts.Count() > 0)
                {
                    handoutMasterID = ++rootID;
                    writeNodeToXML(handoutMasterID, 0, "讲义母版", "", "handoutMaster/", "false");
                    //addRow_Wtree("讲义母版", "andoutMaster", 0, handoutMasterID, "handoutMaster/", "0", "", 1, 9, "0");
                }
                handoutMasterID = rootID;
                int extendedFilePropertiesID = ++rootID;
                writeNodeToXML(extendedFilePropertiesID, 0, "扩展文件属性", "", "extendedFileProperties/", "false");
                //addRow_Wtree("扩展文件属性", "ExtendedFilePropertiesPart", 0, extendedFilePropertiesID, "extendedFileProperties/", "0", "", 1, 10, "0");

                #region Presentation部分
                for (int i = 1; i <= Presentation_parts.ToArray().Length; i++)
                {
                    OpenXmlPart part = ppt.PresentationPart.GetPartById("rId" + i);
                    fileNodeName = part.RootElement.LocalName;
                    #region 幻灯片部分
                    if (fileNodeName == "sld")
                    {
                        c_slides++;
                        xmlFileName = "slide" + c_slides;
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
                            //addRow_Wtree("关联", "rId", CurrentRootID+1, ++rootID, "slide/" + xmlFileName + "/rId", "1", "", 3, 7, "0");
                            for (wi = 1; wi <= p.Parts.Count(); wi++)
                            {
                                writeAttrToXML(++attrID, 0, "rId" + wi, p.GetPartById("rId" + wi).Uri.ToString(), "slide/" + xmlFileName + "/rId" + wi, "0", "0");
                                //addRow_WtreeAttrs("rId" + wi, p.GetPartById("rId" + wi).Uri.ToString(), "slide/" + xmlFileName + "/rId" + wi, "0", "0", 0, 3, 7);
                            }
                        }
                        if (p.SlideLayoutPart.RootElement != null)
                        {
                            writeNodeToXML(++rootID, ++CurrentRootID, "幻灯片版式", "", "slide/" + xmlFileName + "/slideLayout/", "false");
                            //addRow_Wtree("幻灯片版式", "slideLayout", ++CurrentRootID, ++rootID, "slide/" + xmlFileName + "/slideLayout/", "0", "", 3, 7, "0");
                            getAttribute(p.SlideLayoutPart.RootElement, 3, 7, rootID, "slide/" + xmlFileName + "/slideLayout/", 1, null);
                        }
                        if (p.NotesSlidePart != null && p.NotesSlidePart.RootElement != null)
                        {
                            writeNodeToXML(++rootID, CurrentRootID, "备注幻灯片", "", "slide/" + xmlFileName + "/notesSlide/", "false");
                            //addRow_Wtree("备注幻灯片", "notesSlide", CurrentRootID, ++rootID, "slide/" + xmlFileName + "/notesSlide/", "0", "", 3, 8, "0");
                            getAttribute(p.NotesSlidePart.RootElement, 3, 8, rootID, "slide/" + xmlFileName + "/notesSlide/", 1, null);
                        }
                        if (p.SlideCommentsPart != null && p.SlideCommentsPart.RootElement != null)
                        {
                            writeNodeToXML(++rootID, CurrentRootID, "幻灯片批注", "", "slide/" + xmlFileName + "/slideComments/", "false");
                            //addRow_Wtree("幻灯片批注", "slideComment", CurrentRootID, ++rootID, "slide/" + xmlFileName + "/slideComments/", "0", "", 3, 9, "0");
                            getAttribute(p.SlideCommentsPart.RootElement, 3, 9, rootID, "slide/" + xmlFileName + "/slideComments/", 1, null);
                        }
                        if (p.ChartParts.Count() > 0)
                        {
                            int chartPartCount = 0;
                            foreach (ChartPart chartPart in p.ChartParts)
                            {
                                //图标主体
                                if (chartPart.RootElement.LocalName == "chartSpace")
                                {
                                    writeNodeToXML(++rootID, CurrentRootID, "图表空间", "", "slide/" + xmlFileName + "/chartSpace/", "false");
                                    //addRow_Wtree("图表空间", "chartSpace", CurrentRootID, ++rootID, "slide/" + xmlFileName + "/chartSpace/", "0", "", 3, 10, "0");
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
                                    //addRow_Wtree("图表样式", "stylePart", CurrentRootID, ++rootID, "slide/" + xmlFileName + "/chartStyle/", "0", "", 4, (int)hashTable[stylePart.RootElement.ToString()], "0");
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
                                    //addRow_Wtree("图表颜色风格", "colorStylePart", CurrentRootID, ++rootID, "slide/" + xmlFileName + "/chartColorStyle/", "0", "", 4, (int)hashTable[colorStylePart.RootElement.ToString()], "0");
                                    getAttribute(colorStylePart.RootElement, 4, (int)hashTable[colorStylePart.RootElement.ToString()], rootID, "slide/" + xmlFileName + "/chartColorStyle/", (int)hashTable[colorStylePart.RootElement.ToString()], null);
                                }
                                //图表图片
                                if (chartPart.ChartDrawingPart != null && chartPart.ChartDrawingPart.RootElement != null)
                                {
                                    writeNodeToXML(CurrentRootID, ++rootID, get_typeName(chartPart.ChartDrawingPart.RootElement.GetType().ToString()), "", "slide/" + xmlFileName + "/chartSpace/", "true");
                                    //addRow_Wtree(get_typeName(chartPart.ChartDrawingPart.RootElement.GetType().ToString()), get_typeName(chartPart.ChartDrawingPart.RootElement.GetType().ToString()), CurrentRootID, ++rootID, "slide/" + xmlFileName + "/chartSpace/", "1", "", 3, 11, "0");
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
                                //addRow_Wtree("示意图颜色映射", "diagramColorsPart", CurrentRootID, ++rootID, "slide/" + xmlFileName + "/diagramColors/", "0", "", 4, (int)hashTable[diagramColorsPart.RootElement.ToString()], "0");
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
                                //addRow_Wtree("示意图数据", "diagramDataPart", CurrentRootID, ++rootID, "slide/" + xmlFileName + "/diagramData/", "0", "", 4, (int)hashTable[diagramDataPart.RootElement.ToString()], "0");
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
                                //addRow_Wtree("示意图样式", "diagramStylePart", CurrentRootID, ++rootID, "slide/" + xmlFileName + "/diagramStyle/", "0", "", 4, (int)hashTable[diagramStylePart.RootElement.ToString()], "0");
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
                                //addRow_Wtree("示意图保存样式", "diagramPersistLayoutPart", CurrentRootID, ++rootID, "slide/" + xmlFileName + "/diagramPersistLayout/", "0", "", 4, (int)hashTable[diagramPersistLayoutPart.RootElement.ToString()], "0");
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
                                //addRow_Wtree("示意图样式定义", "diagramPersistLayoutPart", CurrentRootID, ++rootID, "slide/" + xmlFileName + "/diagramLayoutDefinition/", "0", "", 4, (int)hashTable[diagramLayoutDefinitionPart.RootElement.ToString()], "0");
                                getAttribute(diagramLayoutDefinitionPart.RootElement, 4, (int)hashTable[diagramLayoutDefinitionPart.RootElement.ToString()], rootID, "slide/" + xmlFileName + "/diagramLayoutDefinition/", (int)hashTable[diagramLayoutDefinitionPart.RootElement.ToString()], null);
                            }
                        }
                    }
                    #endregion

                    else if (fileNodeName == "sldMaster")
                    {
                        c_slideMasters++;
                        xmlFileName = "slideMaster" + c_slideMasters;
                        SlideMasterPart p = (SlideMasterPart)part;
                        if (p.RootElement != null)
                        {
                            int CurrentRootID = rootID + 1;
                            getAttribute(p.RootElement, 2, c_slideMasters, slideMasterID, "slideMaster/" + xmlFileName + "/", c_slideMasters, null);
                            if (p.Parts.Count() > 0)
                            {
                                int wi;
                                writeNodeToXML(++rootID, CurrentRootID, "关联", "", "slideMaster/" + xmlFileName + "/" + "rId", "true");
                                //addRow_Wtree("关联", "slideLayout", CurrentRootID, ++rootID, "slideMaster/" + xmlFileName + "/" + "rId", "1", "", 4, 1, "0");
                                for (wi = 1; wi <= p.Parts.Count(); wi++)
                                {
                                    writeAttrToXML(++attrID, 0, "rId" + wi, p.GetPartById("rId" + wi).Uri.ToString(), "slideMaster/" + xmlFileName + "/" + "rId" + wi, "0", "0");
                                    //addRow_WtreeAttrs("rId" + wi, p.GetPartById("rId" + wi).Uri.ToString(), "slideMaster/" + xmlFileName + "/" + "rId" + wi, "0", "0", 0, 4, 1);
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
                    //addRow_Wtree("关联", "rId", curID, ++rootID, "presentation/rId", "1", "", 2, 1, "0");
                    int wi;
                    for (wi = 1; wi <= ppt.PresentationPart.Presentation.PresentationPart.Parts.Count(); wi++)
                    {
                        writeAttrToXML(++attrID, 0, "rId" + wi, ppt.PresentationPart.Presentation.PresentationPart.GetPartById("rId" + wi).Uri.ToString(), "presentation/rId" + wi, "0", "0");
                        //addRow_WtreeAttrs("rId" + wi, ppt.PresentationPart.Presentation.PresentationPart.GetPartById("rId" + wi).Uri.ToString(), "presentation/rId" + wi, "0", "0", 0, 3, 1);
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
                    String fileName = savePath + paperID + "-" + stuID + "-" + imageIndex + ".gif";
                    img.Save(fileName, System.Drawing.Imaging.ImageFormat.Gif);
                    writeAttrToXML(++attrID, 0, "资源文件", paperID + "-" + stuID + "-" + imageIndex + ".gif", prefix, "0", "0");
                }
                //addRow_Wtree(get_typeName(element.GetType().ToString())+nodeCount, get_typeName(element.GetType().ToString()), fatherID, thisID, prefix, "0", element.InnerText, depth, serial, "0");
                //Console.WriteLine("节点名：{0}\t节点ID：{1}\t父ID：{2}\t深度：{3}\t级：{4}\t前缀：{5}", element.LocalName, thisID, fatherID, depth, serial, prefix);
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
                writeAttrToXML(++attrID, 0, element.LocalName, element.InnerText, prefix, "0", "0");
                //addRow_Wtree(get_typeName(element.GetType().ToString()) + nodeCount, get_typeName(element.GetType().ToString()), fatherID, thisID, prefix, "1", element.InnerText, depth, serial, "0");
                //addRow_WtreeAttrs(element.LocalName, element.InnerText, prefix, "0", "0", 0, depth, serial);
                //Console.WriteLine("节点名：{0}\t文字内容：{1}\t节点ID：{2}\t父ID：{3}\t深度：{4}\t级：{5}\t前缀：{5}", element.LocalName, element.InnerText, thisID, fatherID, depth, serial, prefix);
                return;
            }
            //如果此节点有属性且有子节点
            else if (hasAttributes && hasChildren)
            {
                writeNodeToXML(thisID, fatherID, get_typeName(element.GetType().ToString()) + nodeCount, element.InnerText, prefix, "false");
                //addRow_Wtree(get_typeName(element.GetType().ToString())+nodeCount, get_typeName(element.GetType().ToString()), fatherID, thisID, prefix, "0", element.InnerText, depth, serial, "0");
                //Console.WriteLine("节点名：{0}\t节点ID：{1}\t父ID：{2}\t深度：{3}\t级：{4}\t前缀：{5}", element.LocalName, thisID, fatherID, depth, serial, prefix);
                foreach (OpenXmlAttribute attr in element.GetAttributes())
                {
                    writeAttrToXML(++attrID, 0, attr.LocalName, attr.Value, prefix, "0", "0");
                    //addRow_WtreeAttrs(attr.LocalName, attr.Value, prefix, "0", "0", 0, depth, serial);
                    //Console.WriteLine("节点名：{0}\t属性：{1}\t属性值：{2}\t节点ID：{3}\t父ID：{4}\t深度：{5}\t级：{6}\t前缀：{7}", element.LocalName, attr.LocalName, attr.Value, thisID, fatherID, depth, serial, prefix);
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
                //addRow_Wtree(get_typeName(element.GetType().ToString())+nodeCount, get_typeName(element.GetType().ToString()), fatherID, thisID, prefix, "1", element.InnerText, depth, serial, "0");
                //Console.WriteLine("节点名：{0}\t节点ID：{1}\t父ID：{2}\t深度：{3}\t级：{4}\t前缀：{5}", element.LocalName, thisID, fatherID, depth, serial, prefix);
                foreach (OpenXmlAttribute attr in element.GetAttributes())
                {
                    writeAttrToXML(++attrID, 0, attr.LocalName, attr.Value, prefix, "0", "0");
                    //addRow_WtreeAttrs(attr.LocalName, attr.Value, prefix, "0", "0", 0, depth, serial);
                    //Console.WriteLine("节点名：{0}\t属性：{1}\t属性值：{2}\t节点ID：{3}\t父ID：{4}\t深度；{5}\t级：{6}\t前缀：{7}", element.LocalName, attr.LocalName, attr.Value, thisID, fatherID, depth, serial, prefix);
                }
                return;
            }
        }
        #endregion

        #region 获取英文名
        String get_typeName(String elementType)
        {
            String[] arry = elementType.Split('.');
            String str = arry[arry.Length - 1];//取Document.XX.XX.RunProperties的RunProperties
            String result = "";
            foreach (char c in str)//变成"Run Properties"
            {
                if (c >= 'a' && c <= 'z')//Char.IsLetter(c)小写字母c>='a' && c<='z'
                    result += c;
                else
                    result += " " + c;
            }
            result = result.Substring(0, 1) == " " ? result.Substring(1, result.Length - 1) : result;
            return result;
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
            String score, String status)
        {
            XmlElement element = docAttr.CreateElement("record");
            element.SetAttribute("ID", ID.ToString());
            element.SetAttribute("fid", fatherID.ToString());
            element.SetAttribute("prefix", Prefix);
            element.SetAttribute("attr", attrName);
            element.SetAttribute("value", value);
            element.SetAttribute("score", score);
            element.SetAttribute("status", status);
            element.SetAttribute("paper", paperID.ToString());
            element.SetAttribute("userid", stuID.ToString());
            RootAttr.AppendChild(element);
        }
        #endregion
    }
}
