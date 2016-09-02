/*********** Journal ************/
//Application Indesign Server 
//version 1.0
// Date: 18-10-2012
//Author : Java Team
/// Copyright: Crest-premediaw
//****************************Global  *****************************// 
    //app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;
    logReq = true;
    var startDate = new Date();
    var logMsg = "Writting LOG" + "\n";
    var msgCounter = 0;
    var count = 0;
    var myFloatPlaceFig = "==";
    logMsg = logMsg + "Start Time : " + startDate.toLocaleTimeString() + " Start Date : " + startDate.toDateString() + "\n"
    var myScriptPath = myGetScriptPath();
    var myScriptPath = new Folder (myScriptPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:"));
    var FolderList=myScriptPath.parent;
    var FolderPath = FolderList
    var FolderName = FolderList.toString().split("/").pop(); 
    var myZipPath=File(FolderPath+"/");
    myPGXMLError = myZipPath.getFiles("*.log");
    myPGXMLError = myZipPath.getFiles("*.log");
    temp1 = myPGXMLError.toString().split("/");
    temp = temp1.pop();
    pgTemp = temp.toString().split("_")
    pgRemp1 = pgTemp.pop();
    var Regx=/(Error.log)/gi;
    logMsg = logMsg + (++msgCounter) +". Reading PG xml. " +"\n"
    if(temp.match(Regx))
    {   
        log(myPGXMLError);
        PgXml();
        MoveFolder();
        exit();
    }
    var FolderName = FolderList.toString().split("/").pop();
    var basePath = FolderList.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
    var XmlPath = FolderList+"/"+"OUTXML";
    var PgPath = FolderList+"/"+"Pagination";
    var PdfPath = FolderList+"/"+"PDF";
    var Out = FolderList+"/"+"OUT";
    var zipPath = FolderList;
    var appPath = app.filePath;
    var myControlFile=new File(appPath +"/" +"Configuration"+"/"+ "Config.xml");
    var libraryFilePath=new File(appPath +"/" +"Templates"+"/"+ "Library.indl");
    var myPath="//root/control";
    var figSize=null;
    var journalXmlName="";
    logMsg = logMsg + (++msgCounter) +". Reading Configuration file. " + "\n"
    if(File(myControlFile).exists) 
    {
        myControlFile.open('r');
        var fileString = myControlFile.read();
        var myxml = new XML( fileString );
        journalXmlName=myxml ;
        var Out_Log1=myxml.xpath(myPath)
        myControlFile.close(); 
    }
    else
        alert ("Config.xml  not found");
    var myStyle = "//Zapprocess/Style"
    var myPrefix = "//Zapprocess/Prefix"    
    var LogFileName = "//Zapprocess/GeneralInfo/LogFileName"
    var myTemplate = "//Zapprocess/TemplateName"
    var myColumns = "//Zapprocess/Columns"
    var myJournalID = "//Zapprocess/JournalId"
    var tocRequired = "//Zapprocess/TocRequired"
    var myZipPath=File(FolderList+"/");
    myPGXMLPath = myZipPath.getFiles("*.xml");
    var journalPgXmlName="";//Change made for all at journal dev 
    var myLayoutName="";//Change made for all at journal dev 
    var prefixName="";//Change made for all at journal dev 
    var prefixNumber=0000;//Change made for all at journal dev 
    var myTemplateJournalID =0;
    var myTemplateName ="";
    var isSkapRequired = ""
    var stageID = "";
    var tocRequiredValue  = false;
    var myTemplateColumns="";
        
    /**  Global variable for  writing start and end page number **/
    var _gAddCount = 0;
    var _gStartPageNum = 1;
    var _gEndPageNum = 1;
    
    for (var tp = 0; myPGXMLPath.length>tp; tp++)
    {
        if(File(myPGXMLPath[tp].exists))
        {
            temp1 = myPGXMLPath[tp].toString().split("/");
            temp = temp1.pop();
            tpTemp = temp.toString().split("_")         
            tpRemp1 = tpTemp.pop();
            var Regx=/(_PG.xml)/gi;
            if(temp.match(Regx))
            { 
                var TemplatePath=new File(myPGXMLPath[tp]);
                if(File(TemplatePath).exists) 
                {
                    journalPgXmlName=TemplatePath.name;
                    TemplatePath.open('r');
                    var fileString = TemplatePath.read();
                    var myoutxml = new XML( fileString );
                    var myStyleName=myoutxml.xpath(myStyle)
                    myLayoutName=myStyleName;//Change made for all at journal dev 
                    var myLogFileName=myoutxml.xpath(LogFileName)
                    myTemplateName = myoutxml.xpath(myTemplate)
                    myTemplateColumns = myoutxml.xpath(myColumns)
                    myTemplateJournalID  = myoutxml.xpath(myJournalID)
                    stageID = myoutxml.xpath("//Zapprocess/Workflow/Stageid");
                    myInlineImagePath = myoutxml.xpath("//Zapprocess/Workflow/ImagePath");
                    isSkapRequired = myoutxml.xpath("//Zapprocess/Workflow").@SKAP
                    tocRequiredValue  = myoutxml.xpath(tocRequired)
                    TemplatePath.close();
                    prefixName=myoutxml.xpath(myPrefix )//Change made for all at journal dev 
                    prefixNumber=prefixName.toString().split("_").pop();
                    break;//Change made for all at journal dev 
                }
            }
            else
            {
                log(myPGXMLError);
                PgXml();
                MoveFolder();
                exit();
            }
        }    
    }
    logMsg = logMsg + (++msgCounter) + ". Found Template starting process. " +"\n"
    var myProject=myxml.control 
    var myTemplatepath=null;
    for(var i=0; i<myProject.length(); i++)
    {
        if (myProject[i].@StyleName==myStyleName && myTemplateName!="")
        {
            var colName="";
            colName=getColumnMapValue(myTemplateColumns);
            if(colName!="")
                TemplateName=colName//Change made for all at journal dev 
            tempPath = myProject[i].templatepath+"\\"+myTemplateName+".indt";
            myTemplatepath=tempPath.replace(/\,/g,"\/");
            myDoc = app.open(File(myTemplatepath));
            figSize=myProject[i].@FigSize
        }
    }
    loadScript();
    logMsg = logMsg + (++msgCounter) + ". Detaching all  page elements " + "\n"
    detachAllElementFromPage(myDoc.pages[0]);
    var Pagination= PgPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
    var myfilePath = XmlPath.toString();
    var outXml=myfilePath.replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
    var inPut=File(outXml);
    var inPutFile = inPut.getFiles("*.xml");
    var tempFilename = inPutFile.toString().split("/");
    var temp1Filename = tempFilename.pop();
    var Filenametemp = temp1Filename.toString().split (".xml");
    var Filename = Filenametemp.toString().replace(/\,/g,"");
    var outXml=inPutFile.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
    var myFileFinal = new File(Pagination+"/"+ Filename+".indd");
    var language = "De";
    var SavePath = myFileFinal.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
    var figBasePath = basePath +"/Graphics/PRINT/";
    var equationBasePath = basePath+"/Graphics/Equation/";
    var xmlMetaDataTemplatePathDe=File(appPath+"/"+"Templates/Metadata_ChapterLevel_300_De.indt");
    var xmlMetaDataTemplatePathEn=File(appPath+"/"+"Templates/Metadata_ChapterLevel_300_En.indt");
    var proofVoucherDeTemplatePath=File(appPath+"/"+"Templates/Proof_Procedure_De_Dtitle.indt");
    var proofVoucherEnTemplatePath=File(appPath+"/"+"Templates/Proof_Procedure_En_Dtitle.indt");
    var metaDataAbstractPath=File(appPath+"/"+"Templates/Meta_Temp.indt");
    myDoc.viewPreferences.rulerOrigin=RulerOrigin.pageOrigin;   
    myDoc.viewPreferences.horizontalMeasurementUnits=MeasurementUnits.MILLIMETERS ; 
    myDoc.viewPreferences.verticalMeasurementUnits=MeasurementUnits.MILLIMETERS; 
    myDoc.zeroPoint=[0,0];
    var page=myDoc.pages[0];   
    myDoc.textPreferences.smartTextReflow = true;
    myDoc.textPreferences.limitToMasterTextFrames = false; 
    myDoc.textPreferences.addPages = AddPageOptions.END_OF_DOCUMENT;
    var sAppPath=app.filePath;
    logMsg = logMsg + (++msgCounter) + ". Importing the XML. " + "\n"
    myDoc.xmlElements[0].importXML(outXml);
    try
    {
        logMsg = logMsg + (++msgCounter) + ". Table resizing. " + "\n"
        resizeTableBeforeSave(myDoc);
     }catch(ex){    }
    logMsg = logMsg + (++msgCounter) + ". Starting with Move Elements. " + "\n" 
    MoveElement(journalXmlName, myLayoutName);
    logMsg = logMsg + (++msgCounter) + ". Placing Logo. " +"\n"
    errorFile = File(FolderPath+"/"+ FolderName+"_Error.txt");
    
    findAndPlaceLogo();
    myDoc.save(SavePath); 
    var abstractNodeData="";
    var language = "";
    var myAbstractFile=null;
    var xmlMetaDataFilePath = null;
    var chapterFlie=null;
    var footnoteID = [];
    var temp = myTemplatepath.toString().split("\\");
    var lastImagePlacePage=null;
    var lastTablePlacePage=null;
    var lastBoxPlacePage=null;
    var chapterLanguage=null;
    var topFullWidth=null;
    var topSingleWidth=null;
    var bottomFullWidth=null;
    var bottomSingleWidth=null;
    var lastImagePlacedPage=null;

    try{
        if(logReq == true){
            $.writeln (FolderName+": START: findAndPlaceSocietyLogo()")
            }
        findAndPlaceSocietyLogo();
        if(logReq == true){
            $.writeln (FolderName+": END: findAndPlaceSocietyLogo()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: crossMarkLogoPlacement()")
            }
        crossMarkLogoPlacement();
        if(logReq == true){
            $.writeln (FolderName+": END: crossMarkLogoPlacement()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: crossMarkLogoPlacementHyperlink()")
            }
        crossMarkLogoPlacementHyperlink();
        if(logReq == true){
            $.writeln (FolderName+": END: crossMarkLogoPlacementHyperlink()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: xmp_and_orcid()")
            }
        xmp_and_orcid();
        if(logReq == true){
            $.writeln (FolderName+": END: xmp_and_orcid()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: changeSwatchForLayout()")
            }
        changeSwatchForLayout();
        if(logReq == true){
            $.writeln (FolderName+": END: changeSwatchForLayout()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: changeTableColumnSizeAttribute()")
            }
        //changeTableColumnSizeAttribute();
        if(logReq == true){
            $.writeln (FolderName+": END: changeTableColumnSizeAttribute()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: checkLanguage()")
            }
        checkLanguage();
        if(logReq == true){
            $.writeln (FolderName+": END: checkLanguage()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: changeLanguage()")
            }
        changeLanguage();
        if(logReq == true){
            $.writeln (FolderName+": END: changeLanguage()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: getAbstract()")
            }
        getAbstract();  
        if(logReq == true){
            $.writeln (FolderName+": END: getAbstract()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: tocGeneration()")
            }
        tocGeneration();
        if(logReq == true){
            $.writeln (FolderName+": END: tocGeneration()")
            }
        logMsg = logMsg + (++msgCounter) + ". Getting Chapter language from XML. " + "\n"
        if(logReq == true){
            $.writeln (FolderName+": START: addPBIndentToHere()")
            }
        addPBIndentToHere (myDoc);
        if(logReq == true){
            $.writeln (FolderName+": END: addPBIndentToHere()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: setSuperScriptSubscriptNoBreak()")
            }
        setSuperScriptSubscriptNoBreak(".//Subscript");
        setSuperScriptSubscriptNoBreak(".//Superscript");
        setSuperScriptSubscriptNoBreak(".//cs_text[@type='superscript']");
        if(logReq == true){
            $.writeln (FolderName+": END: setSuperScriptSubscriptNoBreak()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: alignAuthors()")
            }
        alignAuthors();
        if(logReq == true){
            $.writeln (FolderName+": END: alignAuthors()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: alignAuthorsAffiliationEmail()")
            }
        alignAuthorsAffiliationEmail();
        if(logReq == true){
            $.writeln (FolderName+": END: alignAuthorsAffiliationEmail()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: alignKeyword()")
            }
        alignKeyword();
        if(logReq == true){
            $.writeln (FolderName+": END: alignKeyword()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: setEquationSpaceInDefination()")
            }
        setEquationSpaceInDefination();        
        if(logReq == true){
            $.writeln (FolderName+": END: setEquationSpaceInDefination()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: _gRefObj.createHyperlinkCitationRef()")
            }
        _gRefObj.createHyperlinkCitationRef();          
        if(logReq == true){
            $.writeln (FolderName+": END: _gRefObj.createHyperlinkCitationRef()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: _gRefObj.createHyperlinkExternalRef()")
            }
        _gRefObj.createHyperlinkExternalRef();          
        if(logReq == true){
            $.writeln (FolderName+": END: _gRefObj.createHyperlinkExternalRef()")
            }
        myDoc.save(SavePath);
        if(logReq == true){
            $.writeln (FolderName+": START: callLibrary_11812()")
            }
        callLibrary_11812()        
        if(logReq == true){
            $.writeln (FolderName+": END: callLibrary_11812()")
            }
        myDoc.save(SavePath);        
        if(logReq == true){
            $.writeln (FolderName+": START: replaceOperatorSymbols()")
            }
        replaceOperatorSymbols();
        if(logReq == true){
            $.writeln (FolderName+": END: replaceOperatorSymbols()")
            }        
        if(logReq == true){
            $.writeln (FolderName+": START: placeEquation()")
            }
        placeEquation(myDoc);
        if(logReq == true){
            $.writeln (FolderName+": END: placeEquation()")
            }
        myDoc.save(SavePath);            
        if(logReq == true){
            $.writeln (FolderName+": START: addInlineImageInTable()")
            }
        addInlineImageInTable();
        if(logReq == true){
            $.writeln (FolderName+": END: addInlineImageInTable()")
            }
        myDoc.save(SavePath);            
        if(logReq == true){
            $.writeln (FolderName+": START: processTableOptimization()")
            }
         processTableOptimization();
        if(logReq == true){
            $.writeln (FolderName+": END: processTableOptimization()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: createFootNotes()")
            }
        createFootNotes();
        if(logReq == true){
            $.writeln (FolderName+": END: createFootNotes()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: removeWhiteSpaceEquation()")
            }
        removeWhiteSpaceEquation();
        if(logReq == true){
            $.writeln (FolderName+": END: removeWhiteSpaceEquation()")
            }
//~         if(logReq == true){
//~             $.writeln (FolderName+": START: placeEquation()")
//~             }
//~         placeEquation(myDoc);
//~         if(logReq == true){
//~             $.writeln (FolderName+": END: placeEquation()")
//~             }
        if(logReq == true){
            $.writeln (FolderName+": START: placeInlineObject()")
            }
        placeInlineObject();
        if(logReq == true){
            $.writeln (FolderName+": END: placeInlineObject()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: arrange_postcode_affiliation()")
            }
        arrange_postcode_affiliation();
        if(logReq == true){
            $.writeln (FolderName+": END: arrange_postcode_affiliation()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: placeOverviewBox()")
            }
        placeOverviewBox();
        if(logReq == true){
            $.writeln (FolderName+": END: placeOverviewBox()")
            }
        
        if(logReq == true){
            $.writeln (FolderName+": START: addSpaceBetweenTableMergeHead()")
            }
            addSpaceBetweenTableMergeHead();
            addSpaceBetweenTableMergeHead();
            addSpaceBetweenTableMergeHead();
            addSpaceBetweenTableMergeHead();
           applyCaptionStyleIntoTableFirstCell();
           applyCaptionStyleIntoTableFirstCell();
        if(logReq == true){
            $.writeln (FolderName+": END: addSpaceBetweenTableMergeHead()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: applyTableStyle2TopRule()")
            }
        applyTableStyle2TopRule();
        if(logReq == true){
            $.writeln (FolderName+": END: applyTableStyle2TopRule()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: setSuperScriptSubscriptNoBreak()")
            }
        setSuperScriptSubscriptNoBreak(".//Subscript");
        setSuperScriptSubscriptNoBreak(".//Superscript");
        setSuperScriptSubscriptNoBreak(".//cs_text[@type='superscript']");
        if(logReq == true){
            $.writeln (FolderName+": END: setSuperScriptSubscriptNoBreak()")
            }

         myDoc.save(SavePath);    
        if(logReq == true){
            $.writeln (FolderName+": START: equationProcessor()")
            }            
        equationProcessor();            
        if(logReq == true){
            $.writeln (FolderName+": END: equationProcessor()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: floatingPlacement()")
            }
        floatingPlacement();
        if(logReq == true){
            $.writeln (FolderName+": END: floatingPlacement()")
            }
        myDoc.save(SavePath);
        if(logReq == true){
            $.writeln (FolderName+": START: addSpaceBetweenTableMergeHead1()")
            }
        addSpaceBetweenTableMergeHead1();
        if(logReq == true){
            $.writeln (FolderName+": END: addSpaceBetweenTableMergeHead1()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: alignCharTableEntry()")
            }
        alignCharTableEntry();
        if(logReq == true){
            $.writeln (FolderName+": END: alignCharTableEntry()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: applyContinuedTableBottomRule()")
            }
        applyContinuedTableBottomRule();
        if(logReq == true){
            $.writeln (FolderName+": END: applyContinuedTableBottomRule()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: addSidebar()")
            }
        addSidebar(myDoc);
        if(logReq == true){
            $.writeln (FolderName+": END: addSidebar()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: hideSIdebar()")
            }
        hideSIdebar(); 
        if(logReq == true){
            $.writeln (FolderName+": END: hideSIdebar()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: HideFootnote()")
            }
        HideFootnote();
        if(logReq == true){
            $.writeln (FolderName+": END: HideFootnote()")
            }
//~         if(logReq == true){
//~             $.writeln (FolderName+": START: RefreshTableHeadStyles()")
//~             }
//~         RefreshTableHeadStyles(myDoc);
//~         if(logReq == true){
//~             $.writeln (FolderName+": END: RefreshTableHeadStyles()")
//~             }
        if(logReq == true){
            $.writeln (FolderName+": START: refreshFigSideCaptionStyle()")
            }
        refreshFigSideCaptionStyle(myDoc);
        if(logReq == true){
            $.writeln (FolderName+": END: refreshFigSideCaptionStyle()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: addIndentTabInBibiliography()")
            }
        addIndentTabInBibiliography();
        myDoc.save(SavePath);
        if(logReq == true){
            $.writeln (FolderName+": END: addIndentTabInBibiliography()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: AddFrame2FigHyper()")
            }
        AddFrame2FigHyper (myDoc);
        if(logReq == true){
            $.writeln (FolderName+": END: AddFrame2FigHyper()")
            }
        myDoc.save(SavePath);
        if(logReq == true){
            $.writeln (FolderName+": START: _gRefObj.createHyperlinkInternalRef()")
            }
        _gRefObj.createHyperlinkInternalRef();
        if(logReq == true){
            $.writeln (FolderName+": END: _gRefObj.createHyperlinkInternalRef()")
            }
        if(logReq == true){
            $.writeln (FolderName+": START: placeQuestionnaire()")
            }
        placeQuestionnaire(myDoc); 
        if(logReq == true){
            $.writeln (FolderName+": END: placeQuestionnaire()")
            }
        myDoc.save(SavePath);
        if(logReq == true){
            $.writeln (FolderName+": START: refreshFigSideCaptionStyle()")
            }
        refreshFigSideCaptionStyle (myDoc);
        if(logReq == true){
            $.writeln (FolderName+": END: refreshFigSideCaptionStyle()")
            }
    }catch(e){
        alert(e);
        logMsg = logMsg + (++msgCounter) + ". ERROR in first try. " + "\n" + e + "\n"
    }
    finally
    {
        try{
            var rHead=false;
            var dPdf=false; 
            var bookmark=false;
            var pPdf=false;
            myDoc.save(SavePath);
            if(logReq == true){
                $.writeln (FolderName+": START: removeLongHeading()")
                }
            if(myStyleName=="GlobalJournalLarge" || myStyleName=="Medium" || myStyleName=="GlobalJournalSmallCondensed" || myStyleName=="GlobalJournalSmallExtended"){
                removeLongHeading();
            }
            if(logReq == true){
                $.writeln (FolderName+": END: removeLongHeading()")
                }
            if(logReq == true){
                $.writeln (FolderName+": START: MoveElementForRunningHead()")
                }
            MoveElementForRunningHead(myxml, myStyleName);
            if(logReq == true){
                $.writeln (FolderName+": END: MoveElementForRunningHead()")
                }
            myDoc.save(SavePath);
            if(logReq == true){
                $.writeln (FolderName+": START: Deletepdf()")
                }
            Deletepdf();
            if(logReq == true){
                $.writeln (FolderName+": END: Deletepdf()")
                }
            if(logReq == true){
                $.writeln (FolderName+": START: BookMark()")
                }
            BookMark();
            if(logReq == true){
                $.writeln (FolderName+": END: BookMark()")
                }
            if(logReq == true){
                $.writeln (FolderName+": START: hideJournalNodes()")
                }
            hideJournalNodes(journalXmlName, myLayoutName);
            if(logReq == true){
                $.writeln (FolderName+": END: hideJournalNodes()")
                }
            if(logReq == true){
                $.writeln (FolderName+": START: removeBlankLinefromLastPage()")
                }
            removeBlankLinefromLastPage();
            if(logReq == true){
                $.writeln (FolderName+": END: removeBlankLinefromLastPage()")
                }
            if(logReq == true){
                $.writeln (FolderName+": START: tocGeneration()")
                }
            tocGeneration();
            if(logReq == true){
                $.writeln (FolderName+": END: tocGeneration()")
                }
            myDoc.save(SavePath);
            if(logReq == true){
                $.writeln (FolderName+": START: columnReferenceBalance()")
                }
            columnReferenceBalance();
            if(logReq == true){
                $.writeln (FolderName+": END: columnReferenceBalance()")
                }
            myDoc.save(SavePath);
            if(logReq == true){
                $.writeln (FolderName+": START: refresh()")
                }
            refresh();
            if(logReq == true){
                $.writeln (FolderName+": END: refresh()")
                }
            if(logReq == true){
                $.writeln (FolderName+": START: refreshFigSideCaptionStyle()")
                }
            refreshFigSideCaptionStyle(myDoc);
            if(logReq == true){
                $.writeln (FolderName+": END: refreshFigSideCaptionStyle()")
                }
            myDoc.save(SavePath);
            if(logReq == true){
                $.writeln (FolderName+": START: footNoteMarkerPoisiton()")
                }
            footNoteMarkerPoisiton(); 
            if(logReq == true){
                $.writeln (FolderName+": END: footNoteMarkerPoisiton()")
                }
            if(logReq == true){
                $.writeln (FolderName+": START: BiographyImagePlacement()")
                }
           BiographyImagePlacement();
            if(logReq == true){
                $.writeln (FolderName+": END: BiographyImagePlacement()")
                }
            if(logReq == true){
                $.writeln (FolderName+": START: AddQueryReferences()")
                }
            AddQueryReferences(myDoc); 
            if(logReq == true){
                $.writeln (FolderName+": END: AddQueryReferences()")
                }
            if(logReq == true){
                $.writeln (FolderName+": START: deleteEmptyPages()")
                }
            deleteEmptyPages();
            if(logReq == true){
                $.writeln (FolderName+": END: deleteEmptyPages()")
                }
            if(logReq == true){
                $.writeln (FolderName+": START: removeAQEmptyFrame()")
                }
            removeAQEmptyFrame(myDoc);
            if(logReq == true){
                $.writeln (FolderName+": END: removeAQEmptyFrame()")
                }
            myDoc.save(SavePath);
            if(logReq == true){
                $.writeln (FolderName+": START: runSkap()")
                }
            if(isSkapRequired=="true"){
                    runSkap();                
                }
            if(logReq == true){
                $.writeln (FolderName+": END: runSkap()")
                }
            
            if(stageID == "300" && isSkapRequired=="true"){
                }
            else{
            if(logReq == true){
                $.writeln (FolderName+": START: Proofpdf()")
                }
                Proofpdf(myDoc);
            if(logReq == true){
                $.writeln (FolderName+": END: Proofpdf()")
                }
            if(logReq == true){
                $.writeln (FolderName+": START: removeStageInfo()")
                }
                removeStageInfo();
            if(logReq == true){
                $.writeln (FolderName+": END: removeStageInfo()")
                }
            if(logReq == true){
                $.writeln (FolderName+": START: footNoteMarkerPoisiton()")
                }
                footNoteMarkerPoisiton(); 
            if(logReq == true){
                $.writeln (FolderName+": END: footNoteMarkerPoisiton()")
                }
                myDoc.save(SavePath);
            if(logReq == true){
                $.writeln (FolderName+": START: refreshFigSideCaptionStyle()")
                }
                refreshFigSideCaptionStyle(myDoc);
            if(logReq == true){
                $.writeln (FolderName+": END: refreshFigSideCaptionStyle()")
                }
                }
                myDoc.save(SavePath);
                logMsg = logMsg + (++msgCounter) + ". Processing DONE Closing the document. " + "\n"
                
                if(stageID == "300" && isSkapRequired=="true"){
                    }
                else{
                if(logReq == true){
                    $.writeln (FolderName+": START: globalMetaData()")
                    }
                    globalMetaData();
                if(logReq == true){
                    $.writeln (FolderName+": END: globalMetaData()")
                    }
                }
            myDoc.close();
            if(!errorFile.exists){            
                if(logReq == true){
                    $.writeln (FolderName+": START: createLog()")
                    }
                createLog();
                if(logReq == true){
                    $.writeln (FolderName+": END: createLog()")
                    }
            }
            if(logReq == true){
                $.writeln (FolderName+": START: logWritter()")
                }
            logWritter(FolderPath,logMsg);          
            if(logReq == true){
                $.writeln (FolderName+": END: logWritter()")
                }

        }catch(e){
            pdfLog(FolderPath,"Pdf not generated"+e);
            logMsg = logMsg + (++msgCounter) + ". ERROR in FINALLY BLOCK. " + "\n" + e + "\n"
            logWritter(FolderPath,logMsg);
            myDoc.save(SavePath);
            myDoc.close();
            MovePdflog();
            exit();;
        }
    }
function setEquationSpaceInDefination(){
    var mergeHeadEntryNodes=getNodes (myDoc,".//DefinitionList");
    for (st = 0; st < mergeHeadEntryNodes.length; st++){
        applyCorrectStyle(mergeHeadEntryNodes[st])
        }
    }

function setSuperScriptSubscriptNoBreak(nodes){
    var mergeHeadEntryNodes=getNodes (myDoc,nodes);
    for (st = 0; st < mergeHeadEntryNodes.length; st++){
        mergeHeadEntryNodes[st].words.everyItem().noBreak = true;
        }
    }


function applyCorrectStyle(Nodes){
        tabPosition = getEntryTabPosition(Nodes);
        var DifinitinListEntry = Nodes.xmlElements.item("DefinitionListEntry");
        
        for(stb = 0; stb < DifinitinListEntry.index.length; stb++){
        longDifEntry = Nodes.xmlElements[DifinitinListEntry.index[stb]].xmlElements.item("Description").xmlElements.item("Para")
               var appliedParaStyle = longDifEntry.insertionPoints[0].appliedParagraphStyle.name;
               requiredStyle = appliedParaStyle+"_"+(st+1);
               if(!myDoc.paragraphStyles.item(requiredStyle).isValid){
                   duplicateStyle = myDoc.paragraphStyles.item(appliedParaStyle).duplicate();
                   if(duplicateStyle.tabStops.length == 0){
                       duplicateStyle.tabStops.add();
                       }
                   duplicateStyle.tabStops[0].position = tabPosition;
                   duplicateStyle.tabStops[0].alignment = TabStopAlignment.LEFT_ALIGN
                   duplicateStyle.name = requiredStyle;
                   longDifEntry.insertionPoints[0].appliedParagraphStyle = duplicateStyle;
                   longDifEntry.xmlAttributes.item("aid:pstyle").value = requiredStyle;
                   }
               else{
                   longDifEntry.insertionPoints[0].appliedParagraphStyle = myDoc.paragraphStyles.item(requiredStyle);
                  longDifEntry.xmlAttributes.item("aid:pstyle").value = requiredStyle

                   }
            
            }        
        }


function addInlineImageInTable(){
    var inlineImageInTableNodes=getNodes (myDoc,".//entry//InlineMediaObject/ImageObject");
    $.writeln (myInlineImagePath)
    for (st = 0; st < inlineImageInTableNodes.length; st++){
        if(inlineImageInTableNodes[st].xmlAttributes.item("FileRef").isValid){
            inlineImageName = inlineImageInTableNodes[st].xmlAttributes.item("FileRef").value;
            inlineImageName = inlineImageName.split(".")[0];
            myInlineImagePath = "D:\\Breeze-pagination\\Breeze\\IN\\Work\\10488_2015_646_Article\\Graphics\\WEB\\";
           myInlineImage = File(myInlineImagePath + inlineImageName +".eps");
            if(!myInlineImage.exists)
            {
                myInlineImage=File("D:/FigNotFound.jpg");
            }
            var Insp = inlineImageInTableNodes[st].insertionPoints[0];
            rect = Insp.place(myInlineImage);
            rect [0].parent.appliedObjectStyle = myDoc.objectStyles.item("InlineEqFrameObjectStyle");

        }
        }
    }



function getEntryTabPosition(Nodes){
    var tabPosition = null;
    var max_cont_entry = getMaximumEntryTermContent(Nodes)
    longDifEntry = Nodes.xmlElements[max_cont_entry];
    if(longDifEntry.xmlElements[1].markupTag.name=="cs_text"){
        if(longDifEntry.xmlElements[1].xmlAttributes.item("type").isValid){
            if(longDifEntry.xmlElements[1].xmlAttributes.item("type").value=="vbtab"){
               longDifEntry.xmlElements[1].contents = SpecialCharacters.EM_SPACE;
               tabFrame = longDifEntry.xmlElements.item ("Description").insertionPoints[0].textFrames.add();
               tabPosition = tabFrame.geometricBounds[1];
               parentTextFrame = longDifEntry.xmlElements.item ("Description").insertionPoints[0].parentTextFrames[0];
               parentTextFramePos = parentTextFrame.geometricBounds[1];
               tabPosition = tabPosition - parentTextFramePos;
               tabFrame.remove();
               longDifEntry.xmlElements[1].contents = "	"
               }
            }
        }
    return tabPosition;
    }
function applyTableStyle2TopRule(){
    var style2TableNodes=getNodes (myDoc,".//table[@Float='Yes' and @Style='Style2']//Table");
    for(st = 0; st < style2TableNodes.length; st++){
        myFirstRow = style2TableNodes[st].tables[0].rows[1];
        for (stR = 0; stR < myFirstRow.cells.length; stR++){
            if(myFirstRow.cells[stR].appliedCellStyle.name=="Table_Head_Left"){
                if(!myFirstRow.cells[stR].associatedXMLElement.xmlAttributes.item("aid5:cellstyle").isValid){
                    myFirstRow.cells[stR].associatedXMLElement.xmlAttributes.add("aid5:cellstyle", "Table_Head_Left_2/3")
                    }
                myFirstRow.cells[stR].associatedXMLElement.xmlAttributes.item("aid5:cellstyle").value = "Table_Head_Left_2/3"
                myFirstRow.cells[stR].appliedCellStyle = myDoc.cellStyles.item("Table_Head_Left_2/3");
                }
            else if(myFirstRow.cells[stR].appliedCellStyle.name=="clTableHead"){
                if(!myFirstRow.cells[stR].associatedXMLElement.xmlAttributes.item("aid5:cellstyle").isValid){
                    myFirstRow.cells[stR].associatedXMLElement.xmlAttributes.add("aid5:cellstyle", "clTableHead_2/3")
                    }
                myFirstRow.cells[stR].appliedCellStyle = myDoc.cellStyles.item("clTableHead_2/3");
                }
            if(myFirstRow.cells[stR].appliedCellStyle.name=="Table_Head_Right"){
                if(!myFirstRow.cells[stR].associatedXMLElement.xmlAttributes.item("aid5:cellstyle").isValid){
                    myFirstRow.cells[stR].associatedXMLElement.xmlAttributes.add("aid5:cellstyle", "Table_Head_Right_2/3")
                    }
                myFirstRow.cells[stR].appliedCellStyle = myDoc.cellStyles.item("Table_Head_Right_2/3")
                }
            }
        }
    }

function getMaximumEntryTermContent(Nodes){
    var DifinitinListEntry = Nodes.xmlElements.item("DefinitionListEntry");
    max_label_data = 0;
    defaultData = 0
    for(stb = 0; stb < DifinitinListEntry.index.length; stb++){
        defTermContents = Nodes.xmlElements[DifinitinListEntry.index[stb]].xmlElements.item("Term").contents;
        if(defTermContents.length > defaultData){
            defaultData = defTermContents.length;            
            max_label_data = stb;
            }
        }
    return max_label_data;
    }

function addSpaceBetweenTableMergeHead(){
    var mergeHeadEntryNodes=getNodes (myDoc,".//entry[@nameend]");
    var tableFrame = null
    var allTextFrames=app.documents[0].pages.lastItem().textFrames;
    for(var t=0;t<allTextFrames.length;t++)
    {
        if(allTextFrames[t].associatedXMLElement.markupTag.name =="Article")
            tableFrame= allTextFrames[t];
            break;
    }

    myCurrentTextSize = tableFrame.geometricBounds
    myCurrentTextSize1 = myCurrentTextSize;
    tableFrame.geometricBounds = [myCurrentTextSize[0], myCurrentTextSize[1], myCurrentTextSize[0]+100000, myCurrentTextSize[3]]
    
    for (stOA = 0; stOA < mergeHeadEntryNodes.length; stOA++){
        
            if(mergeHeadEntryNodes[stOA].parent.cells[0].parentRow.rowType==1162375799){
            //myTableFrame = mergeHeadEntryNodes[stOA].insertionPoints[0].parentTextFrames[0]
            var entryColoumnPos = mergeHeadEntryNodes[stOA].xmlAttributes.item("nameend").value;
            entryColoumnPos = entryColoumnPos.substring (1, entryColoumnPos.length);
            entryColumnNo = mergeHeadEntryNodes[stOA].parent.cells[0].parentColumn.parent.columns.count()
            headingNo = countHeaderRow(mergeHeadEntryNodes[stOA].parent.cells[0].parent)
            if(headingNo > 2){
                if(headingNo-1 > mergeHeadEntryNodes[stOA].parent.cells[0].rowSpan){
                    if(entryColoumnPos < entryColumnNo){
                        if(mergeHeadEntryNodes[stOA].parent.cells[0].parentRow.index==1){
                            mergeHeadEntryNodes[stOA].parent.cells[0].appliedCellStyle = myDoc.cellStyles.item("clTableHead_Bridge");
                            mergeHeadEntryNodes[stOA].applyCellStyle (myDoc.cellStyles.item("clTableHead_Bridge"));
                            if(!mergeHeadEntryNodes[stOA].parent.xmlAttributes.item("aid5:cellstyle").isValid){
                                mergeHeadEntryNodes[stOA].parent.xmlAttributes.add ("aid5:cellstyle", "clTableHead_Bridge");
                                }
                                
                            mergeHeadEntryNodes[stOA].parent.xmlAttributes.item("aid5:cellstyle").value = "clTableHead_Bridge";
                            //mergeHeadEntryNodes[stOA].parent.cells[0].bottomEdgeStrokeColor = myDoc.swatches.item("None");
                            //mergeHeadEntryNodes[stOA].parent.cells[0].clearCellStyleOverrides (true);
                            }
                        else{
                            mergeHeadEntryNodes[stOA].parent.cells[0].appliedCellStyle = myDoc.cellStyles.item("clTableHead_Bridge_1");
                            mergeHeadEntryNodes[stOA].applyCellStyle (myDoc.cellStyles.item("clTableHead_Bridge_1"));
                            if(!mergeHeadEntryNodes[stOA].parent.xmlAttributes.item("aid5:cellstyle").isValid){
                                mergeHeadEntryNodes[stOA].parent.xmlAttributes.add ("aid5:cellstyle", "clTableHead_Bridge_1");
                                }
                            mergeHeadEntryNodes[stOA].parent.xmlAttributes.item("aid5:cellstyle").value = "clTableHead_Bridge_1";
                            }
                        if(mergeHeadEntryNodes[stOA].paragraphs.length > 1){
                            if(mergeHeadEntryNodes[stOA].xmlElements.lastItem().markupTag.name == "SimplePara"){
                                mergeHeadEntryNodes[stOA].xmlElements.lastItem().applyParagraphStyle(myDoc.paragraphStyles.item("Table_Head_Bridge_Rule"));            
                                if(!mergeHeadEntryNodes[stOA].xmlElements.lastItem().xmlAttributes.item("aid:pstyle").isValid){
                                    mergeHeadEntryNodes[stOA].xmlElements.lastItem().xmlAttributes.add ("aid:pstyle", "Table_Head_Bridge_Rule");
                                    }
                                else{
                                    mergeHeadEntryNodes[stOA].xmlElements.lastItem().xmlAttributes.item("aid:pstyle").value = "Table_Head_Bridge_Rule";
                                    }
                                }
                            }
                        else{
                                mergeHeadEntryNodes[stOA].applyParagraphStyle(myDoc.paragraphStyles.item("Table_Head_Bridge_Rule"));            
                                mergeHeadEntryNodes[stOA].xmlAttributes.item("aid:pstyle").value = "Table_Head_Bridge_Rule";                            
                            }
                        }
                    else if(entryColoumnPos == entryColumnNo){
                        if(mergeHeadEntryNodes[stOA].paragraphs.length > 1){
                            if(mergeHeadEntryNodes[stOA].xmlElements.lastItem().markupTag.name == "SimplePara"){
                                mergeHeadEntryNodes[stOA].xmlElements.lastItem().applyParagraphStyle(myDoc.paragraphStyles.item("Table_Head_Bridge_Rule"));            
                                if(!mergeHeadEntryNodes[stOA].xmlElements.lastItem().xmlAttributes.item("aid:pstyle").isValid){
                                    mergeHeadEntryNodes[stOA].xmlElements.lastItem().xmlAttributes.add ("aid:pstyle", "Table_Head_Bridge_Rule");
                                    }
                                else{
                                    mergeHeadEntryNodes[stOA].xmlElements.lastItem().xmlAttributes.item("aid:pstyle").value = "Table_Head_Bridge_Rule";
                                    }
                                }
                            }
                        else{
                                mergeHeadEntryNodes[stOA].applyParagraphStyle(myDoc.paragraphStyles.item("Table_Head_Bridge_Rule"));            
                                mergeHeadEntryNodes[stOA].xmlAttributes.item("aid:pstyle").value = "Table_Head_Bridge_Rule";                            
                            }
                        }
                    }
                }
            //myTableFrame.fit(FitOptions.FRAME_TO_CONTENT);
            }
        }
        tableFrame.geometricBounds = myCurrentTextSize1
    }


function addSpaceBetweenTableMergeHead1(){
    var mergeHeadEntryNodes=getNodes (myDoc,".//entry[@nameend]");
    for (stOA = 0; stOA < mergeHeadEntryNodes.length; stOA++){
        
            if(mergeHeadEntryNodes[stOA].parent.cells[0].parentRow.rowType==1162375799){
            newEnter = getLine(mergeHeadEntryNodes[stOA].parent.cells[0]); 
            if (newEnter > 0){
                for (inL = 0; inL < newEnter; inL++){
                        mergeHeadEntryNodes[stOA].insertTextAsContent ("\n", XMLElementPosition.ELEMENT_END);
                    }
                }
            }
        }
    }

function getLine(myCell){
    requiredLine = 0;
    myDoc.viewPreferences.horizontalMeasurementUnits=MeasurementUnits.POINTS;
    myDoc.viewPreferences.verticalMeasurementUnits=MeasurementUnits.POINTS;
    cellLine = myCell.lines.length;
    cellHeight = myCell.height;
    cellTopMargin = myCell.topInset;
    cellBottomMargin = myCell.bottomInset;
    cellHeight = cellHeight - (cellBottomMargin);
    cellLeading = myCell.insertionPoints[0].leading;
    currentLine = cellHeight / cellLeading;
    currentLine = currentLine.toFixed ();
    if(currentLine > cellLine){
        requiredLine = currentLine - cellLine;
        }
    myDoc.viewPreferences.horizontalMeasurementUnits=MeasurementUnits.MILLIMETERS ; 
    myDoc.viewPreferences.verticalMeasurementUnits=MeasurementUnits.MILLIMETERS; 
    
    return requiredLine;     
    }

function applyCaptionStyleIntoTableFirstCell(){
    var tableNodes=getNodes (myDoc,".//table[@Float='Yes']/Story/Table");
    for(i = 0; i < tableNodes.length; i++){
        if(tableNodes[i].xmlElements[0].cells[0].appliedCellStyle.name=="clTableCaption"){
            tableNodes[i].xmlElements[0].cells[0].appliedCellStyle = myDoc.cellStyles.item("clTableCaption");
            tableNodes[i].xmlElements[0].cells[0].clearCellStyleOverrides()
            }
        }
    }

    function countHeaderRow(aTable){
        contH = 0;
        for (i = 0; i < aTable.rows.count(); i++){
            if(aTable.rows[i].rowType==1162375799){
                contH = contH + 1;
                }
            }
        return contH;
        }
    
function alignCharTableEntry(){
var charEntryNodes=getNodes (myDoc,".//entry[@align='char']");
for (stOA = 0; stOA < charEntryNodes.length; stOA++){
    try{
    myTableFrame = charEntryNodes[stOA].insertionPoints[0].parentTextFrames[0]
    charEntryVal = charEntryNodes[stOA].xmlAttributes.item("char").value;
    charEntryType = charEntryValueConvertion();
    tableID = charEntryNodes[stOA].parent.parent.parent.parent.xmlAttributes.item("ID").value
    entryColumnNo = charEntryNodes[stOA].parent.cells[0].parentColumn.index;
    cellWidth = charEntryNodes[stOA].parent.cells[0].width;
    cellLeftMargin = charEntryNodes[stOA].parent.cells[0].leftInset;
    cellRightMargin = charEntryNodes[stOA].parent.cells[0].rightInset;
//~     actualCellSpace = cellWidth - (cellLeftMargin+cellRightMargin);
    entryApStyleName = "Table_Body_"+tableID+"_"+charEntryType+"_"+entryColumnNo;
    reqStyleName = myDoc.paragraphStyles.item(entryApStyleName);
    anchorFirstTxtFrm = charEntryNodes[stOA].insertionPoints[0].textFrames.add();
    with(anchorFirstTxtFrm.anchoredObjectSettings){
        anchoredPosition = AnchorPosition.ANCHORED;
        horizontalReferencePoint = AnchoredRelativeTo.ANCHOR_LOCATION;
        anchorPoint = AnchorPoint.BOTTOM_LEFT_ANCHOR;
        }
    textFirstPos = anchorFirstTxtFrm.geometricBounds[1];
    anchorFirstTxtFrm.remove();
//    $.writeln ()
    anchorEndTxtFrm = charEntryNodes[stOA].insertionPoints[charEntryNodes[stOA].contents.indexOf(charEntryVal)].textFrames.add();
    with(anchorEndTxtFrm.anchoredObjectSettings){
        anchoredPosition = AnchorPosition.ANCHORED;
        horizontalReferencePoint = AnchoredRelativeTo.ANCHOR_LOCATION;
        anchorPoint = AnchorPoint.BOTTOM_LEFT_ANCHOR;
        }
    textLastPos = anchorEndTxtFrm.geometricBounds[1];
    anchorEndTxtFrm.remove();
    
    actualCellSpace = (textLastPos - textFirstPos);
    
    if(!reqStyleName.isValid){
        myDupStyle = myDoc.paragraphStyles.item("Table_Body").duplicate();
        myDupStyle.firstLineIndent = 0;
        myDupStyle.leftIndent = 0;
        myDupStyle.name = entryApStyleName;
        t = myDupStyle.tabStops.add();
        t.alignment = TabStopAlignment.CHARACTER_ALIGN
        if(charEntryVal=="−"){
           charEntryVal = "-"
            }
        t.alignmentCharacter = charEntryVal;
        t.position = actualCellSpace;
        }
    else{
        if(actualCellSpace > reqStyleName.tabStops[0].position){
            reqStyleName.tabStops[0].position = actualCellSpace;
            }
        }
    charEntryNodes[stOA].applyParagraphStyle (reqStyleName);
    charEntryNodes[stOA].xmlAttributes.item("aid:pstyle").value = entryApStyleName;       
    myTableFrame.fit(FitOptions.FRAME_TO_CONTENT);
        }//try end
    catch(e){}
    }
}



function charEntryValueConvertion(){
    charValStr = "";
    switch(charEntryVal){
        case '=':
            charValStr = "Equal";
            break;
        case '−':
            charValStr = "Minus";
            break;
        case '+':
            charValStr = "Plus";
            break;
        case '±':
            charValStr = "PlusMinus";
            break;
        case '–':
            charValStr = "EnDash";
            break;
        case '.':
        case '·':
            charValStr = "Decimal";
            break;
        case '×':
            charValStr = "Multiply";
            break;
        case ',':
            charValStr = "Comma";
            break;
        case '(':
            charValStr = "OpenBracket";
            break;
//~         case '.':
//~             charValStr = "MidDot";
//            break;
       default:
            charValStr = charEntryVal;                   
        }
    return charValStr;
    }

    
    function arrange_postcode_affiliation(){
        var OrgAddress=getNodes (myDoc,".//test_OrgAddress");
        for (stOA = 0; stOA < OrgAddress.length; stOA++){
            for(stOAC = 0; stOAC < OrgAddress[stOA].xmlElements.length-2; stOAC++){
                currElement = OrgAddress[stOA].xmlElements;
                if(currElement.item(stOAC).markupTag.name=="test_Postcode" && currElement.item(currElement.item(stOAC).index+2).markupTag.name=="test_City"){
                    if(currElement.item(stOAC).insertionPoints[0].baseline!=currElement.item(currElement.item(stOAC).index+2).insertionPoints[-1].baseline){
                        currElement.item(stOAC).insertTextAsContent ("\n", XMLElementPosition.ELEMENT_START)
                        }
                    }
                else if(currElement.item(stOAC).markupTag.name=="test_City" && currElement.item(currElement.item(stOAC).index+2).markupTag.name=="test_Postcode"){
                    if(currElement.item(stOAC).insertionPoints[0].baseline!=currElement.item(currElement.item(stOAC).index+2).insertionPoints[-1].baseline){
                        currElement.item(stOAC).insertTextAsContent ("\n", XMLElementPosition.ELEMENT_START)
                        }
                    }
                else if(currElement.item(stOAC).markupTag.name=="test_State" && currElement.item(currElement.item(stOAC).index+2).markupTag.name=="test_Postcode"){
                    if(currElement.item(stOAC).insertionPoints[0].baseline!=currElement.item(currElement.item(stOAC).index+2).insertionPoints[-1].baseline){
                        currElement.item(stOAC).insertTextAsContent ("\n", XMLElementPosition.ELEMENT_START)
                        }
                    }
                else if(currElement.item(stOAC).markupTag.name=="test_Postcode" && currElement.item(currElement.item(stOAC).index+2).markupTag.name=="test_State"){
                    if(currElement.item(stOAC).insertionPoints[0].baseline!=currElement.item(currElement.item(stOAC).index+2).insertionPoints[-1].baseline){
                        currElement.item(stOAC).insertTextAsContent ("\n", XMLElementPosition.ELEMENT_START)
                        }
                    }
                else if(currElement.item(stOAC).markupTag.name=="test_Postcode" && currElement.item(currElement.item(stOAC).index+3).markupTag.name=="test_Country"){
                    if(currElement.item(stOAC).insertionPoints[0].baseline!=currElement.item(currElement.item(stOAC).index+3).insertionPoints[-1].baseline){
                        currElement.item(stOAC).insertTextAsContent ("\n", XMLElementPosition.ELEMENT_START)
                        }
                    }
                }
            }                    
        }

    function removeBlankLinefromLastPage(){
        myParentElementStory = myDoc.xmlElements[0].insertionPoints[-1].parentStory;
        myBodyStoryParagraph = myParentElementStory.paragraphs;

        for(i = myBodyStoryParagraph.length-1; i >= 0; i--){
            //myBodyStoryParagraph[i].select()
//~             $.writeln (myBodyStoryParagraph[i].contents)
            if(myBodyStoryParagraph[myBodyStoryParagraph.length-1].words.length!=0 && myBodyStoryParagraph[myBodyStoryParagraph.length-1].appliedParagraphStyle.name.indexOf ("Literatur_Entry")!=-1){
                break;
                }
            
            lastCharAssociatedElement = myBodyStoryParagraph[i].characters.lastItem().insertionPoints[-1].associatedXMLElements[0];
            if(lastCharAssociatedElement.markupTag.name=="cs_text"){
                if(lastCharAssociatedElement.xmlAttributes.item("type").isValid){
                    if(lastCharAssociatedElement.xmlAttributes.item("type").value=="vbnewline"){
                        if(myBodyStoryParagraph[i].words.length==0 && myBodyStoryParagraph[i+1].words.length > 0){
                            myBodyStoryParagraph[i].appliedParagraphStyle = myBodyStoryParagraph[i+1].appliedParagraphStyle       
                            lastCharAssociatedElement.contents = "";
                            myBodyStoryParagraph[i].clearOverrides ();
                            }
                        else if(myBodyStoryParagraph[i].words.length > 0 && myBodyStoryParagraph[i+1].words.length == 0){
                            myBodyStoryParagraph[i+1].appliedParagraphStyle = myBodyStoryParagraph[i].appliedParagraphStyle       
                            lastCharAssociatedElement.contents = "";
                            myBodyStoryParagraph[i].clearOverrides ();
                            }
                        else if(myBodyStoryParagraph[i].words.length==0 && myBodyStoryParagraph[i-1].words.length==0){
                            myBodyStoryParagraph[i].appliedParagraphStyle = myBodyStoryParagraph[i-1].appliedParagraphStyle       
                            lastCharAssociatedElement.contents = "";
                            myBodyStoryParagraph[i].clearOverrides ();
                            }
                        else if(myBodyStoryParagraph[i].words.length==0 && myBodyStoryParagraph[i-1].words.length > 0){
                            myBodyStoryParagraph[i].appliedParagraphStyle = myBodyStoryParagraph[i-1].appliedParagraphStyle       
                            lastCharAssociatedElement.contents = "";
                            myBodyStoryParagraph[i].clearOverrides ();
                            }
                        else if(myBodyStoryParagraph[i].appliedParagraphStyle.name.indexOf ("Literatur_Entry")!=-1){
                            break;
                            }
                        }
                    }
                }        
            }
        }
    
    function alignKeyword(){
        var keywordNodes=getNodes (myDoc, ".//KeywordGroup/Keyword");
        for(var l=0;l<keywordNodes.length;l++)
        {
            var keywordNode=keywordNodes[l];
            if(keywordNode.insertionPoints[0].baseline!=keywordNode.insertionPoints[-1].baseline){
                keywordNode.insertTextAsContent ("\n", XMLElementPosition.ELEMENT_START)
                }
        }
    }

    function alignAuthors(){
        var AuthorNameNodes=getNodes (myDoc, ".//cs_text[@type='authname']/AuthorName");
        for(var l=AuthorNameNodes.length-1;l>=0;l--)
        {
            var AuthorNameNode=AuthorNameNodes[l];
            AuthorNameNode.parent.insertionPoints[0].parentTextFrames[0].fit(FitOptions.FRAME_TO_CONTENT);
            if(AuthorNameNode.insertionPoints[0].baseline!=AuthorNameNode.xmlElements.item("FamilyName").insertionPoints[-1].baseline){
                AuthorNameNode.insertTextAsContent ("\n", XMLElementPosition.ELEMENT_START);
                AuthorNameNode.insertionPoints[0].parentTextFrames[0].fit(FitOptions.FRAME_TO_CONTENT);
                }
        }
    }

    function alignAuthorsAffiliationEmail(){
        var AuthorNameNodes=getNodes (myDoc, ".//cs_text[@type='auth_aff_collections']//test_Email");
        for(var l=0;l<AuthorNameNodes.length;l++)
        {
            var AuthorNameNode=AuthorNameNodes[l];
            if(AuthorNameNode.insertionPoints[0].baseline!=AuthorNameNode.insertionPoints[-1].baseline){
                AuthorNameNode.insertTextAsContent ("\n", XMLElementPosition.ELEMENT_START);
                AuthorNameNode.insertionPoints[0].parentTextFrames[0].fit(FitOptions.FRAME_TO_CONTENT);
                }
        }
    }


    function applyContinuedTableBottomRule(){
            var tableNodes = getNodes(myDoc, ".//table[@Float='Yes']/Story/Table");
            for(var i=0;i<tableNodes.length;i++)
            {
                var count = 0; 
                tableNode = tableNodes[i];
                if(tableNode.tables[0].columns[0].cells[0].insertionPoints[0].parentTextFrames[0].nextTextFrame!=null)
                {
                    columns_n = tableNode.tables[0].columns;
                    for(st = 0; st < columns_n.count(); st++){
                        for(ii = 0; ii < columns_n[st].cells.count() -1; ii++)
                        {
                            cellObj  = columns_n[st].cells[ii]
                            currentCell = cellObj.insertionPoints[0].parentTextFrames[0];
                            rowId = cellObj.insertionPoints[0].associatedXMLElements[0].xmlElements.item("entry").xmlAttributes.item("cs_rowindex");
                            moreRows = cellObj.insertionPoints[0].associatedXMLElements[0].xmlElements.item("entry").xmlAttributes.item("morerows");
                            findNextRow = columns_n[st].cells[ii+1].insertionPoints[0].associatedXMLElements[0].xmlElements.item("entry").xmlAttributes.item("cs_rowindex");
                            
//~                             findNextRow = columns_n[st].parent.rows.item(cellObj.parentRow.index+1);
                            nextCell = columns_n[st].cells[ii+1].insertionPoints[0].parentTextFrames[0];    
                            if(findNextRow.isValid)
                            {
                                nextRowId = findNextRow;
                                if (rowId.isValid ==true && nextRowId.isValid==true)
                                {
                                    nextRowColumnIndex = columns_n[st].parent.rows.item(columns_n[st].cells[ii].parentRow.index).cells[0].parentColumn.index;
                                    if(moreRows.isValid){
                                    currentRowValue = parseInt(rowId.value) + parseInt(moreRows.value);
                                        }
                                    else{
                                    currentRowValue = rowId.value;
                                    }
                                    nextRowValue = nextRowId.value
                                    if(parseInt(currentCell.parentPage.name)!=parseInt(nextCell.parentPage.name) && (currentRowValue ==(nextRowValue -1)))
                                    {
                                    cellObj.bottomEdgeStrokeWeight = "0.5 pt";
                                    cellObj.bottomEdgeStrokeColor = myDoc.swatches.item("Black");
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    

    function removeLongHeading(RunningXPath){
        var running_title_Nodes=getNodes (myDoc, ".//cs_text[@type='right_running_title']");
        for(var l=0;l<running_title_Nodes.length;l++)
        {
            var running_title_Node=running_title_Nodes[l];
            traverseRunningHeadingNode(running_title_Node);
        }
    }

    function traverseRunningHeadingNode(running_title_Node){
       // running_title_Node.applyParagraphStyle (myDoc.paragraphStyles.item ("Running_Head_Recto"), false)
        current_Page = myDoc.pages.lastItem();
        leftMargin = current_Page.marginPreferences.left;
        rightMargin = current_Page.marginPreferences.right;
        documentWidth = myDoc.documentPreferences.pageWidth;
        textWidth = documentWidth - (leftMargin+rightMargin);
        dummyHeadFrame = current_Page.textFrames.add();
        dummyHeadFrame.geometricBounds = [0, 0, current_Page.marginPreferences.top, textWidth]
        dummyHeadFrame.placeXML (running_title_Node);
        dummyHeadFrame.insertionPoints[0].applyParagraphStyle (myDoc.paragraphStyles.item ("Running_Head_Recto"), false);
        if(dummyHeadFrame.lines.count() > 1){
            removeMoreLinesRunningHead();
            }
        dummyHeadFrame.fit(FitOptions.FRAME_TO_CONTENT);
        heading_width = dummyHeadFrame.geometricBounds[3]-dummyHeadFrame.geometricBounds[1];
        removeExtraRunningHead();        
        }
    function removeMoreLinesRunningHead(){
        for(stln = 1; stln < dummyHeadFrame.lines.count(); stln++){
            dummyHeadFrame.lines[stln].remove();
            }
        }
    function removeExtraRunningHead(){
        if(myStyleName=="GlobalJournalLarge" || myStyleName=="Medium"){
            running_head_width = 129;
            }
        else if(myStyleName=="GlobalJournalSmallCondensed"){
            running_head_width = 80;
            }
        else if(myStyleName=="GlobalJournalSmallExtended"){
            running_head_width = 78.1;
            }

        if(heading_width < running_head_width){
            dummyHeadFrame.remove();
            }
        else{
            while(heading_width > running_head_width){
                dummyHeadFrame.words.lastItem().remove();
                dummyHeadFrame.characters.lastItem().remove();

                dummyHeadFrame.fit(FitOptions.FRAME_TO_CONTENT);
                heading_width = dummyHeadFrame.geometricBounds[3]-dummyHeadFrame.geometricBounds[1];                
                }
            if(dummyHeadFrame.contents.substring(dummyHeadFrame.contents.length-1, dummyHeadFrame.contents.length)==" "){
            dummyHeadFrame.insertionPoints[-1].contents = ". . .";
                }
            else{
            dummyHeadFrame.insertionPoints[-1].contents = " . . .";
                }
            dummyHeadFrame.remove();
            }    
        }

    function placeOverviewBox(){
        if(myTemplateName=="BSL_Journal_T1_Grey"){
            //var overViewNode=getNodes (myDoc,".//*[RenderingStyle='Style1' and @Type='Overview']");
            //var overViewNode="//*[RenderingStyle='Style1' and @Type='Overview']";
            var xpath ="//*[@Type='Overview']";
            var root  = myDoc.xmlElements[0];
            var node  = null;
            try{
                var proc  = app.xmlRuleProcessors.add([xpath]);
                var match = proc.startProcessingRuleSet(root);
                while( match!=undefined ) 
                {
                    node = match.element;                    
                    match = proc.findNextMatch();
                    precedding_style = node.insertionPoints[-1].appliedParagraphStyle.name
                    node.insertTextAsContent ("\r", XMLElementPosition.AFTER_ELEMENT)
//~                     if(precedding_style=="Head_1" || precedding_style=="Head_2" || precedding_style=="Head_3" || precedding_style=="Head_4" || precedding_style=="Head_5" || precedding_style=="Head_1_No above rule" || precedding_style=="Head_2_After 1" || precedding_style=="Head_3_After 2" || precedding_style=="Head_4_After 3" || precedding_style=="Head_5_After 4"){
//~                         node.insertTextAsContent ("\r", XMLElementPosition.AFTER_ELEMENT)
//~                         }
                    var overFrame=node.placeIntoInlineFrame([117, 10]);
                    refresh();
                    overFrame.appliedObjectStyle=myDoc.objectStyles.item("ProcedureBK"); 
                    overFrame_pos = overFrame.geometricBounds;
                    overFrame.geometricBounds = [overFrame.geometricBounds[0], overFrame.geometricBounds[1], overFrame.geometricBounds[0]+172, overFrame.geometricBounds[3]]
                    var docWidth=myDoc.documentPreferences.pageWidth;
                    var docHeight=myDoc.documentPreferences.pageHeight;                 
                    overFrame.fit(FitOptions.FRAME_TO_CONTENT)

//~                     if(overFrame.overflows==false){
//~                         overFrame.fit(FitOptions.FRAME_TO_CONTENT)
//~                         }
//~                     else{
//~                         var top=overFrame.parentPage.marginPreferences.top;
//~                         var bottom=overFrame.parentPage.marginPreferences.bottom;
//~                         var left=overFrame.parentPage.marginPreferences.left;
//~                         var right=overFrame.parentPage.marginPreferences.right;
//~                         overFrameFloat = overFrame.parentPage.textFrames.add();
//~                         overFrameFloat.geometricBounds = [overFrame_pos[0], left, docHeight-bottom, docWidth-right]
//~                         overFrameFloat.appliedObjectStyle=myDoc.objectStyles.item("ProcedureBK"); 
//~                         node.placeXML (overFrameFloat);
//~                         overFrame.remove();  
//~                         while(overFrameFloat.parentStory.overflows){
//~                             myPage = myDoc.pages.item(parseInt (overFrameFloat.parentPage.name, 10))
//~                             var top=myPage.marginPreferences.top;
//~                             var bottom=myPage.marginPreferences.bottom;
//~                             var left=myPage.marginPreferences.left;
//~                             var right=myPage.marginPreferences.right;                            
//~                             overFrameFloat_1 = myPage.textFrames.add();
//~                             overFrameFloat_1.geometricBounds = [top, left, docHeight-bottom, docWidth-right]
//~                             overFrameFloat_1.appliedObjectStyle=myDoc.objectStyles.item("ProcedureBK"); 
//~                             overFrameFloat_1.previousTextFrame = overFrameFloat;
//~                             overFrameFloat = overFrameFloat_1;
//~                             }
//~                         overFrameFloat.fit(FitOptions.FRAME_TO_CONTENT)
//~                         }
                }
            }catch( ex ){
                alert(ex);
            }finally{
                proc.endProcessingRuleSet();
                proc.remove();
                }
            }
        }


    function callLibrary_11812()
    {
        if(myTemplateJournalID=="11812"){
            if(myDoc.xmlElements[0].markupTag.name=="Article"){
                if(myDoc.xmlElements[0].xmlElements.item("cs_tocnode")!=null){
                    cs_tocnode_index = myDoc.xmlElements[0].xmlElements.item("cs_tocnode").index
                    myElement = myDoc.xmlElements[0].xmlElements.item(cs_tocnode_index)
                    myLib = app.open (File (File(myTemplatepath).parent+"/CME_DFP.indl"))
                    var myLibObj = myLib.assets[0].placeAsset (myElement.insertionPoints[-1])                
                    myLib.close()
                    myDoc.xmlElements[0].xmlElements.item(cs_tocnode_index+1).insertTextAsContent (SpecialCharacters.PAGE_BREAK, XMLElementPosition.ELEMENT_START)
                    }
                }
            }
    }


    function processTableOptimization()
    {
        var myScriptPath = File (sAppPath+"/"+"Scripts/Breeze/Table_optimization.jsxbin");
        myScriptPath = myScriptPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        Fpath=File(myScriptPath);
        myRefresh = Fpath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        app.doScript(myRefresh);
    }

    function xmp_and_orcid()
    {
        var myScriptPath = File (sAppPath+"/"+"Scripts/Breeze/SKAP/XMPandORCID.jsxbin");
        myScriptPath = myScriptPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        Fpath=File(myScriptPath);
        myRefresh = Fpath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        app.doScript(myRefresh);
    }

    function runSkap()
    {
        var myScriptPath = File (sAppPath+"/"+"Scripts/Breeze/SKAPPAGINATION.jsxbin");
        myScriptPath = myScriptPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        Fpath=File(myScriptPath);
        myRefresh = Fpath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        app.doScript(myRefresh);
    }


    function removeStageInfo()
    {
        var myScriptPath = File (sAppPath+"/"+"Scripts/Breeze/SKAP/RemoveStageInfo.jsxbin");
        myScriptPath = myScriptPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        Fpath=File(myScriptPath);
        myRefresh = Fpath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        app.doScript(myRefresh);
    }

    function equationProcessor()
    {
        var myScriptPath = File (sAppPath+"/"+"Scripts/Breeze/EpsProcessor.jsxbin");
        myScriptPath = myScriptPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        Fpath=File(myScriptPath);
        myRefresh = Fpath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        app.doScript(myRefresh);
    }

    function adjustAuthorCollectionFrame()
    {
        logMsg = logMsg + (++msgCounter) + ". Adjust author collection frame. " + "\n"
        if(TemplateName=="2_Col_Journals.indt")
        {
            var myPage = myDoc.pages[0]
            for (var j = 0; myPage.pageItems.length>j;  j++) 
            {
                var mypg = myPage.pageItems[j]
                try{
                    if(mypg.label=="auth_aff_collections")
                    {
                        var y1 = mypg.geometricBounds[0];
                        var x1 = mypg.geometricBounds[1];
                        var w = mypg.geometricBounds[3];
                        var h = mypg.geometricBounds[2];
                        if((Math.round(h-y1))>120)
                        {
                                          
                          var textFrameWidth=myDoc.documentPreferences.pageWidth-myPage.marginPreferences.left-myPage.marginPreferences.right;
                            var docWidth=myDoc.documentPreferences.pageWidth;
                            var docHeight=myDoc.documentPreferences.pageHeight;                 
                            var top=myPage.marginPreferences.top;
                            var bottom=myPage.marginPreferences.bottom;
                            var left=myPage.marginPreferences.left;
                            var right=myPage.marginPreferences.right;
                            mypg.geometricBounds=[y1,x1,docHeight-bottom,textFrameWidth+x1];
                            mypg.textFramePreferences.verticalJustification=VerticalJustification.TOP_ALIGN;
                            mypg.fit(FitOptions.FRAME_TO_CONTENT);
                            mypg.appliedObjectStyle = myDoc.objectStyles.item("TextFrame_AuthorDetails_2_col");
                            mypg.fit(FitOptions.FRAME_TO_CONTENT);
                            mypg.geometricBounds=[mypg.geometricBounds[0],mypg.geometricBounds[1],mypg.geometricBounds[2]+5,mypg.geometricBounds[3]]
                            var textColumn=maxLineColumn(mypg);
                            var val=textColumn.lines.lastItem().baseline +1;
                            mypg.geometricBounds=[mypg.geometricBounds[0],mypg.geometricBounds[1],val,mypg.geometricBounds[3]];
                            mypg.textFramePreferences.verticalJustification=VerticalJustification.BOTTOM_ALIGN;
                            var docHeight=myDoc.documentPreferences.pageHeight;                 
                            var bottom=myDoc.pages[0].marginPreferences.bottom;
                            var getDifference=docHeight-bottom-mypg.geometricBounds[2];
                            mypg.move([mypg.geometricBounds[1],mypg.geometricBounds[0]+getDifference])
                        }
                    }
                } catch(e){}
            }
        }
    }

    function hideJournalNodes(myxml, myStyleName)
    {
        logMsg = logMsg + (++msgCounter) + ". Hiding nodes. " + "\n"
        var myProject=myxml.control;
        var myPageItems = myDoc.allPageItems;
        var myPage = myDoc.pages.everyItem();
        for(var i=0; i<myProject.length(); i++)
        {
            if(myProject[i].@StyleName==myStyleName)
            {
                var Xml_Log1=myProject[i].Element.*
                for (var e=0; e<Xml_Log1.length(); e++)
                {
                    var myElement = Xml_Log1[e].@Element;
                    var myAttributes = Xml_Log1[e].@type;
                    var hideAttributes = Xml_Log1[e].@hide;
                    var xpath ="//"+myElement+"[@"+"type"+"="+"'"+ myAttributes+"'"+"]";
                    var root  = myDoc.xmlElements[0];
                    var node  = null;
                    if(hideAttributes==true)
                        cs_text_hide(myAttributes)
                }
            }
        }
    }

    function cs_text_hide(a)
    {
        var xPath = "//cs_text[@type='"+a.toString()+"']";
        var myDoc= app.documents.item(0);
        var root  = myDoc.xmlElements[0];
        var node  = null;
        try{
            var proc  = app.xmlRuleProcessors.add([xPath]);
            var match = proc.startProcessingRuleSet(root);
            while(match!=undefined) 
            {
                node = match.element;
                match = proc.findNextMatch();
                content=node.contents;
                { 
                    if(node!=null && node!=undefined)
                    {
                        if(node.xmlAttributes.itemByName("type").isValid)
                        {
                            var text = "Hide";
                            var content = node.contents;
                            var tagName = node.markupTag.name.toString ()
                            var Hide = node.xmlInstructions.item(0);
                            var string = '<cs_text>' + content +'</cs_text>';
                            var myXMLProcessingInstruction = node.xmlInstructions.add(text, string);
                            Hide.move(LocationOptions.before, node);
                            try
                            {
                                node.remove();
                            }catch(e){}
                        }
                    }
                }
            }
        }
        catch( ex ){
            alert(ex);
        }
        finally{
            proc.endProcessingRuleSet();
            proc.remove();
        }
    }

    function applyStyleToInlineEquation()
    {
        var inlineEquations=getNodes (myDoc, ".//InlineEquation");
        for(var l=0;l<inlineEquations.length;l++)
        {
            var inlineEqaution=inlineEquations[l];
        }
    }

    function footNoteMarkerPoisiton()
    {
        logMsg = logMsg + (++msgCounter) + ". Footnote marker position. " + "\n"
        myDoc.footnoteOptions.markerPositioning=FootnoteMarkerPositioning.NORMAL_MARKER;
        myDoc.save(SavePath);
        myDoc.footnoteOptions.markerPositioning=FootnoteMarkerPositioning.SUPERSCRIPT_MARKER;
    }

    function deleteEmptyPages()
    {
        logMsg = logMsg + (++msgCounter) + ". Delete Blank Pages. " + "\n"
        var pages = app.documents[0].pages.everyItem().getElements();
        for(var i = pages.length-1;i>=0;i--)
        {
            var removePage = true;
            if(pages[i].pageItems.length>0)
            {
                var items = pages[i].pageItems.everyItem().getElements();
                for(var j=0;j<items.length;j++)
                {
                    if(items[j].label.toLowerCase()=="table"||items[j].label.toLowerCase().indexOf("tab")==0)
                    {
                        removePage=false;
                        break;
                    }
                    if(items[j].constructor.name.toLowerCase()=="rectangle")
                    {
                        removePage=false;
                        break;
                    }
                     if(!(items[j] instanceof TextFrame))
                    {
                        if((items[j].constructor.name.toLowerCase()=="rectangle") && (TemplateName=="2_Col_Journals.indt" || TemplateName=="1_Col_Journals.indt"))
                            continue;
                        else
                        {
                            removePage=false;
                            break;
                        }
                    }
                    if((items[j].contents!=""&&items[j].label=="") || (items[j].contents!=""&&items[j].label.toLowerCase()=="main3") || (items[j].contents!=""&&items[j].label.toLowerCase()=="main2") || (items[j].contents!=""&&items[j].label.toLowerCase()=="main1"))
                    {
                        removePage=false;
                        break;
                    }
                }
            }
            if(i==0 && app.documents[0].pages.length==1)
                removePage = false
            if((removePage&&pages[i].appliedMaster.name=="A-Normal Page") || (removePage&&pages[i].appliedMaster.name=="G-Index") || (removePage&&pages[i].appliedMaster.name=="B-Master"))
                pages[i].remove()
        }
    }

    
function placeQuestionnaire(myDoc)
{
    try{
      if(myTemplateName=="Vienna_11812_Template")
     { 
        var questionnaireNode=getNodes (myDoc,".//*[(@Type='Questionnaire')]");
        var questionnaireStartPage =questionnaireNode[0].insertionPoints[0].parentTextFrames[0].parentPage;
        var questionnaireEndPage =questionnaireNode[0].insertionPoints[-1].parentTextFrames[0].parentPage;
        var questionnaireTextFrame =questionnaireNode[0].insertionPoints[0].parentTextFrames[0];
        var questNextTextFrame=null;
        var newQuestionnaireEndPage=null;

        for(var i=parseInt (questionnaireStartPage.name)-1;i<parseInt (questionnaireEndPage.name);i++)
        {
            var topPrevious=myDoc.pages[i].marginPreferences.top;
            var bootomPrevious=myDoc.pages[i].marginPreferences.bottom;
            var bottomDifference=null;
            var topDifference=null;    
            
             
                if( myDoc.pages[i].name==questionnaireStartPage.name)
                {
                    myDoc.pages[i].appliedMaster=myDoc.masterSpreads.item("E-BM_Master");
                }
                else
                {
                      myDoc.pages[i].appliedMaster=myDoc.masterSpreads.item("F-BM_Text");
                }
                questionnaireTextFrame.appliedObjectStyle = myDoc.objectStyles.item("WB_QuestionsFrame");
           
            var left= myDoc.pages[i].marginPreferences.left;
            var right= myDoc.pages[i].marginPreferences.right;
            var top= myDoc.pages[i].marginPreferences.top;
            var bottom= myDoc.pages[i].marginPreferences.bottom;
            var textFrameWidth=myDoc.documentPreferences.pageWidth-left-right;
            var width=textFrameWidth-(questionnaireTextFrame.geometricBounds[3]-questionnaireTextFrame.geometricBounds[1])    
            questNextTextFrame=questionnaireTextFrame.nextTextFrame;
            var bottomDifference=bottom-bootomPrevious;
            var topDifference=top-topPrevious;
            
              if(questionnaireTextFrame!=null )
              {
                    if(parseInt(myDoc.pages[i].name)%2==0)
                    {
                        myDoc.align([questionnaireTextFrame],AlignOptions.LEFT_EDGES,AlignDistributeBounds.MARGIN_BOUNDS)
                    }
                    
                        questionnaireTextFrame.geometricBounds=[questionnaireTextFrame.geometricBounds[0]+topDifference,questionnaireTextFrame.geometricBounds[1],questionnaireTextFrame.geometricBounds[2]-bottomDifference,questionnaireTextFrame.geometricBounds[3]+width];
                    if(parseInt (questionnaireEndPage.name)>=parseInt(questionnaireTextFrame.parentPage.name))
                    {
                        questionnaireTextFrame=questNextTextFrame;
                     }   
              
              }
             newQuestionnaireEndPage=questionnaireNode[0].insertionPoints[-1].parentTextFrames[0].parentPage;
             questionnaireEndPage=newQuestionnaireEndPage;
        }
   }
    }catch(e){}

}
    
    function RefreshTableHeadStyles(myDoc)
    {
        logMsg = logMsg + (++msgCounter) + ". Refresh table Head Styles. " + "\n"
        var allCells=getNodes(myDoc, ".//Cell") 
        for(var i=0;i<allCells.length;i++)
        {
            var entryAttr=allCells[i];
            if(entryAttr.xmlAttributes.item("aid5:cellstyle").isValid)
            {
                try{
                    if(entryAttr.xmlAttributes.item("aid5:cellstyle").value=="clTableHead")
                    {
                        var idInssPt  = entryAttr.insertionPoints.item(0);
                        idInssPt.parent.appliedCellStyle=myDoc.cellStyles.item("clTableHead");
                        var tableFrame=idInssPt.parent.parent.parent;
                        if(tableFrame instanceof TextFrame)
                            tableFrame.fit(FitOptions.frameToContent);
                    }
                    if(entryAttr.xmlAttributes.item("aid5:cellstyle").value=="clTableCaption")
                    {
                        var idInssPt  = entryAttr.insertionPoints.item(0);
                        idInssPt.parent.appliedCellStyle=myDoc.cellStyles.item("clTableCaption");
                        var tableFrame=idInssPt.parent.parent.parent;
                        if(tableFrame instanceof TextFrame)
                            tableFrame.fit(FitOptions.frameToContent);
                    }
                }catch(e){ }
            }            
        }
    }
    
    function addPBIndentToHere(myDoc)
    {
        
        var ReqNds = getNodes(myDoc,".//cs_text[@type='indentohere' or  @type='FMpageBreak' or @type='MayBeFMPageBreak' or @type='MayBeBMPageBreak']|.//cs_text[@type='BMpageBreak']|.//cs_text[@type='softenter']|.//cs_text[@type='RightIndentTab'] | .//cs_text[@type='mspace']");
        for(var  i = 0;i< ReqNds.length;i++)
        {
            var Nd= ReqNds[i];
            if (Nd.xmlAttributes.item("type").value == "FMpageBreak") 
                Nd.contents = SpecialCharacters.pageBreak;
            if (Nd.xmlAttributes.item("type").value == "RightIndentTab") 
                Nd.contents = SpecialCharacters.RIGHT_INDENT_TAB;
            if (Nd.xmlAttributes.item("type").value == "mspace") 
                Nd.contents = SpecialCharacters.EM_SPACE;
            if (Nd.xmlAttributes.item("type").value == "BMpageBreak") 
                Nd.contents = SpecialCharacters.pageBreak;
            if (Nd.xmlAttributes.item("type").value == "indentohere") 
                Nd.contents = SpecialCharacters.INDENT_HERE_TAB;
        }
    }

    function addIndentTabInBibiliography()
    {
        logMsg = logMsg + (++msgCounter) + ". Add tab indent to Biblography. " + "\n"
        var ReqNds = getNodes(myDoc,".//cs_text[@type='indentohere']");
        for(var  i = 0;i< ReqNds.length;i++)
        {
            var Nd= ReqNds[i];
            if (Nd.xmlAttributes.item("type").value == "indentohere") 
            {
                var parentNode=Nd.parent.parent;
                if(parentNode.markupTag.name.toLowerCase()=="citation")
                    Nd.contents = SpecialCharacters.INDENT_HERE_TAB;
                var headingNode=Nd.parent;
                if(headingNode.markupTag.name.toLowerCase()=="heading")
                    Nd.contents = SpecialCharacters.INDENT_HERE_TAB;        
            }
        }
    }
    
    function refresh()
    {
        logMsg = logMsg + (++msgCounter) + ". Refreshing Document. " + "\n"
        var myScriptPath = File (sAppPath+"/"+"Scripts/Breeze/Refresh.jsxbin");
        myScriptPath = myScriptPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        Fpath=File(myScriptPath);
        myRefresh = Fpath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        app.doScript(myRefresh);
        setSuperScriptSubscriptNoBreak(".//Subscript");
        setSuperScriptSubscriptNoBreak(".//Superscript");
        setSuperScriptSubscriptNoBreak(".//cs_text[@type='superscript']");
        replaceOperatorSymbols();
    }
    
    function addSidebar(myDoc)
    {
        logMsg = logMsg + (++msgCounter) + ". Add Side bar. " + "\n"
        var allparagraphNodes = getNodes(myDoc,".//cs:reposition");
        for(var i=0; i<allparagraphNodes.length;i++)
        {
             var paragraphNode = allparagraphNodes[i].xmlElements;
             for(var j=0;j<paragraphNode.length;j++)
             {
                if(paragraphNode[j].markupTag.name.toLowerCase()=="sidebar")
                {
                    var repoidPara  = allparagraphNodes[i].xmlAttributes.item("cs:repoID").value;
                    var sidebarParagraph = getNodes(myDoc, ".//cs_repos[@repoID='" + repoidPara + "']")[0] ;
                    var sideFrame=sidebarParagraph.placeIntoInlineFrame([44, 30]);
                    sideFrame.appliedObjectStyle=myDoc.objectStyles.item("Sidebar");
                    myDoc.save();
                    var pageName=sidebarParagraph.insertionPoints[0].parentTextFrames[0].parentPage.name;
                    var contents=paragraphNode[j].texts[0];
                    contents.duplicate(LocationOptions.AT_BEGINNING,sideFrame) ; 
                    sideFrame.fit(FitOptions.frameToContent);
                }
            }
        }
    }
    
     function  hideSIdebar()
    {
        logMsg = logMsg + (++msgCounter) + ". Hide Side bar. " + "\n"
        var myScriptPath = File (sAppPath+"/"+"Scripts/Breeze/Hide-Sidebar-Journal.jsxbin");
        myScriptPath = myScriptPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        Fpath=File(myScriptPath);
        myRefresh = Fpath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        app.doScript(myRefresh);
    }

    function HideFootnote()
    {
        logMsg = logMsg + (++msgCounter) + ". Hiding Footnotes. " + "\n"
        var xPath = "//footnote";
        var myDoc= app.documents.item(0);
        var root  = myDoc.xmlElements[0];
        var node  = null;
        try{
            var proc  = app.xmlRuleProcessors.add([xPath]);
            var match = proc.startProcessingRuleSet(root);
            while(match!=undefined) 
            {
                node = match.element;
                match = proc.findNextMatch();
                content=node.contents; 
                if(node!=null && node!=undefined && node.xmlAttributes.itemByName("cs_type").value!="endnote")
                {            
                    var text = "Hide";
                    var content = node.contents;
                    var tagName = node.markupTag.name.toString ()
                    var type = node.xmlAttributes.itemByName("cs_type").value;
                    if(type.toString()!="endnote")
                    {
                        var Hide = node.xmlInstructions.item(0);
                        var string = '<' +tagName.toString()+" " +"cs_type" +"="+'"'+type.toString()+'"'+'>'+content.toString() +'</'+tagName.toString()+'>'
                        var myXMLProcessingInstruction = node.xmlInstructions.add(text, string);
                        Hide.move(LocationOptions.before, node);
                    }
                    try{
                        node.remove();
                    }catch(e){}                        
                }
            }
        }catch( ex ){
            alert(ex);
        }
    }
    
    function align_hor( obj, source_align, target_align )
    {
        var target = getPosition( obj[0], target_align )
        for( var i = 1; i < obj.length; i++ )
        obj[i].move( undefined, 
        [target - getPosition(obj[i],source_align), 0], true )
    }

    function getPosition( o, pos )
    {
        switch( pos )
        {
            case 'hcentre' : return o.geometricBounds[1] + ((o.geometricBounds[3] - o.geometricBounds[1]) / 2);
        }
    }

    function getWidthofBiggestPageObject(page,textframe,eyeCatcher)
    {
        var object=null;
        var height=null;
        var pageObject=page.allPageItems;
        for(var  i=0;i<pageObject.length;i++)
        {
            var test=false;
            if(textframe!=undefined)
                test=(pageObject[i].id==textframe.id);
             if((pageObject[i].label.toLowerCase()=="table" ||pageObject[i].label.indexOf("AnchorFrame")!=-1)&&!test)
            {
                var tableWidth=pageObject[i].geometricBounds[3]-pageObject[i].geometricBounds[1];
                var textFrameWidth=((((myDoc.documentPreferences.pageWidth-page.marginPreferences.left)-page.marginPreferences.right)-page.marginPreferences.columnGutter)/2);
                if(Math.round(tableWidth)>Math.round(textFrameWidth))
                {
                    if(height==null)
                    {
                        height=pageObject[i].geometricBounds[2];
                        object=pageObject[i];
                    }
                    if(height!=null&&pageObject[i].geometricBounds[2]>height)  
                    {
                        height=pageObject[i].geometricBounds[2];
                        object=pageObject[i];
                    }
                }
            }
            if(pageObject[i].constructor.name.toLowerCase()=="group" && pageObject[i].label!="TipEyecatcherBox")
            {
                var figWidth=pageObject[i].geometricBounds[3]-pageObject[i].geometricBounds[1];
                var textFrameWidth=((((myDoc.documentPreferences.pageWidth-page.marginPreferences.left)-page.marginPreferences.right)-page.marginPreferences.columnGutter)/2);
                if(Math.round(figWidth)>Math.round(textFrameWidth)&&eyeCatcher==undefined)
                {
                    if(height==null)
                    {
                        height=pageObject[i].geometricBounds[2];
                        object=pageObject[i];
                    }
                    if(height!=null&&pageObject[i].geometricBounds[2]>height)  
                    {
                        height=pageObject[i].geometricBounds[2];
                        object=pageObject[i];
                    }
                }
                else if((Math.round(figWidth-1)<=Math.round(textFrameWidth)||Math.round(figWidth-1)>Math.round(textFrameWidth))&&eyeCatcher==true )
                {
                    var tipTextFrame=textframe.previousTextFrame;  
                    var tipGroup=tipTextFrame.parent;
                    var tipParent= pageObject[i].parent;
                    if(tipParent.constructor.name.toLowerCase()=="group")
                        continue;
                    if(pageObject[i].id==tipGroup.id)
                        continue;
                    else
                    {
                        if(height==null)
                        {
                            height=pageObject[i].geometricBounds[2];
                            object=pageObject[i];
                        }
                        if(height!=null&&pageObject[i].geometricBounds[2]>height)  
                        {
                            height=pageObject[i].geometricBounds[2];
                            object=pageObject[i];
                        }
                    }
                }
            }
        }
        return object;
    }

    function nextPage(parentPage)
    {
        var nextPageName=parseInt(parentPage.name)+1;
        var nextPage=myDoc.pages[nextPageName-1] ;    
        return nextPage; 
    }

    function newNextPage(object)
    {
        var page=null;
        var textFrame=findPlacedTextFrame (object);
        var nextPage=null;
        var figureNextTextFrame=textFrame.nextTextFrame;
        if(figureNextTextFrame.parentPage.marginPreferences.columnCount==1)
            nextPage=figureNextTextFrame.parentPage;
        else
        {
            while(parseInt(figureNextTextFrame.parentPage.name)!=parseInt(figureNextTextFrame.nextTextFrame.parentPage.name))
            {
                nextPage=figureNextTextFrame.parentPage;
                figureNextTextFrame=figureNextTextFrame.nextTextFrame;
            }
            nextPage=figureNextTextFrame.parentPage;
        }
        if(nextPage!=null ||nextPage!=undefined)
            page=nextPage;
        return page;
    }
    
    function getSingleNodeValue(expression,startingNode)
    {
        var xPath = expression;
        var root  = startingNode;
        var node  = null;
        try{
            var proc  = app.xmlRuleProcessors.add([xPath, ]);
            var match = proc.startProcessingRuleSet(root);
            while( match!=undefined ) 
            { 
                node = match.element;
                match = proc.findNextMatch(); 
            }
        }catch( ex ){
            alert(ex);
        }
        finally{
            proc.endProcessingRuleSet();
            proc.remove();
        }
        if(node!=null)
            return node;
        else
            return "";
    }    
    
    function lastObjectPlacePage()
    {
        var imagePage=null;
        var tablePage=null;
        var largest =null;
        var boxPage=null;
        var array=new Array();
        if(lastImagePlacePage!=null)
        {
            imagePage=parseInt(lastImagePlacePage.name);
            array.push(imagePage);
        }
        if(lastTablePlacePage!=null)
        {
            tablePage=parseInt(lastTablePlacePage.name);
            array.push(tablePage);
        }
        if(lastBoxPlacePage!=null)
        {
            boxPage=parseInt(lastBoxPlacePage.name);
            array.push(boxPage);
        }
        if(array.length>0)
        {
            var array = [imagePage, tablePage, boxPage];
            largest = Math.max.apply(Math, array);
        }   
        return largest;
    }
    
    function getNodes(doc, expression) 
    {
        var data= new Array();
        try
        {
            data= doc.xmlElements[0].evaluateXPathExpression(expression,[['cs' , 'http://www.crest-premedia.in'],['aid', 'http://ns.adobe.com/AdobeInDesign/4.0/']]);
        }catch(e){}
        return data;
    }

    function alert1()
    {
    }

    function getPageByObject (_object) 
    {
        if (_object != null) 
        {
            _object = _object.getElements ()[0]; // Problems with
            if (_object.hasOwnProperty("baseline")) 
                _object = _object.parentTextFrames[0];
            while (_object != null) 
            {
                if (_object.hasOwnProperty ("parentPage")) 
                    return _object.parentPage;
                var whatIsIt = _object.constructor;
                switch (whatIsIt) 
                {
                    case Page : return _object;
                    case Character : _object = _object.parentTextFrames[0]; break;
                    case Footnote :; // drop through
                    case Cell : _object = _object.insertionPoints[0].parentTextFrames[0]; break;
                    case Note : _object = _object.storyOffset.parentTextFrames[0]; break;
                    case XMLElement : if (_object.insertionPoints[0] != null) { _object = _object.insertionPoints[0].parentTextFrames[0]; break; }
                    case Application : return null;
                    default: _object = _object.parent;
                }
                if (_object == null) 
                    return null;
            }
            return _object;	
        } 
        else 
            return null;
    }

    function placeEquation(myDoc)
    {
        logMsg = logMsg + (++msgCounter) + ". Placing Equations. " + "\n"
        myDoc.save(SavePath); 
        var allEquations  = getNodes(myDoc,".//*[@cs_type='Equation']");    
        for(var i  = 0;i< allEquations.length;i++)
        {
            try{
                var eq = allEquations[i];
                if (eq.xmlAttributes.item("cs_category").value == "Numbered") 
                    DispUnnumEquation(myDoc,eq);
                else if (eq.xmlAttributes.item("cs_category").value == "Unnumbered") 
                    DispUnnumEquation(myDoc,eq);
                else
                    InlineEquation(myDoc,eq);
                var style = applyStyletoEquation();
            }catch(e){
                alert("equation not placed="+e);
            }
        }
    }

    function DispUnnumEquation(myDoc,EqNode)
    {
        var EqPath  = "";
        var Insp  = EqNode.insertionPoints[0];
        var myLeading = EqNode.insertionPoints[0].appliedParagraphStyle.leading
        var IndObj = null;
        EqPath = EqNode.xmlAttributes.item("cs_EqPath").value;
        if (EqPath != ""){
            IndObj = Insp.place(new File(EqPath));
            }
        else{
            IndObj = Insp.place(new File("c:\EquationMissing.eps"));
            }
        var rex = IndObj[0].parent;
        rex.appliedObjectStyle = myDoc.objectStyles.item("EquationFrameObjectStyle");
        if(EqPath != ""){
            IndObj[0].itemLink.unlink()
        }
        try{
            equationFrmSZ = rex.geometricBounds;
            equationHeight = equationFrmSZ[2]-equationFrmSZ[0]
        }
        catch(err){
            equationFrmSZ = checkEquationHeight(EqPath)
            equationHeight = equationFrmSZ * 25.4
            }
        if(myDoc.viewPreferences.verticalMeasurementUnits==2054188905/*PT*/){
            equationHeight = equationHeight;        
            }
        else if(myDoc.viewPreferences.verticalMeasurementUnits==2053991795/*MM*/){
            equationHeight = equationHeight*2.83;                
            }
        myLine = convertLine(equationHeight, myLeading);
        xmlPos = checkXMLPosition(EqNode);
        eqStyle = "EQ_"+myLine+"_"+xmlPos;
        //$.writeln (eqStyle)
        Insp.applyParagraphStyle(myDoc.paragraphStyles.item(eqStyle))
        if(EqNode.markupTag.name=="Equation"){
        if(EqNode.xmlAttributes.item("aid:pstyle").isValid){
            if(EqNode.xmlAttributes.item("aid:pstyle").value=="EQ"){
                EqNode.xmlAttributes.item("aid:pstyle").value = eqStyle;
                }
            }
        }        
        myDoc.save(SavePath);
    }
    
    function convertLine(eqHeight, leadingOfNode){
        LINE_CONV = eqHeight/leadingOfNode
        LINE_NO = LINE_CONV.toFixed(0)
//~         $.writeln ("Height: "+eqHeight)
//~         $.writeln ("Leading: "+leadingOfNode)
//~         $.writeln ("LINE_CONV: "+LINE_CONV)
//~         $.writeln ("Line No: "+LINE_NO)
        return LINE_NO;
        }
    function checkEquationHeight(equationPath){
        eqFileObj = File(equationPath)
        eqFileObj.open('r')
    do{
        line = eqFileObj.readln();
        if(line.substring (0, 13) == "%%Dimensions:"){
            line_cont = line.split (" ")
            line_cont_H = line_cont[1].split ("=")[1]
            }
        }
    while (!myFile.eof) {
    }
    myFile.close()
        }
    function checkXMLPosition(xmlElementNode){
        //xmlElementNode.select();
        equationType = "";
        parentElement = xmlElementNode
        grandParentElement = parentElement.parent
        xml_element_position = parentElement.index
        if(xml_element_position==1){
            if(grandParentElement.xmlElements.item (xml_element_position+2).isValid){
                if(grandParentElement.xmlElements.item (xml_element_position+2).markupTag.name=="Equation"){
                    return equationType = "FIRST"
                    }
                else{
                    return equationType = "ONLY"
                    }                
                }
            else{
                return equationType = "ONLY"
                }
            }
        else{
            if( grandParentElement.xmlElements.item (xml_element_position-2).isValid && 
                grandParentElement.xmlElements.item (xml_element_position+2).isValid){
                if( grandParentElement.xmlElements.item (xml_element_position-2).markupTag.name=="Equation" &&
                    grandParentElement.xmlElements.item (xml_element_position+2).markupTag.name=="Equation"){
                        return equationType = "MID";                        
                    }
                else if( grandParentElement.xmlElements.item (xml_element_position-2).markupTag.name!="Equation" &&
                    grandParentElement.xmlElements.item (xml_element_position+2).markupTag.name!="Equation"){
                        return equationType = "ONLY";                        
                    }                    
                else if( grandParentElement.xmlElements.item (xml_element_position-2).markupTag.name!="Equation" &&
                    grandParentElement.xmlElements.item (xml_element_position+2).markupTag.name=="Equation"){
                        return equationType = "FIRST";                        
                    }                    
                }
            else if( grandParentElement.xmlElements.item (xml_element_position-2).isValid && 
                !grandParentElement.xmlElements.item (xml_element_position+2).isValid){
                if( grandParentElement.xmlElements.item (xml_element_position-2).markupTag.name=="Equation"){
                        return equationType = "LAST"                        
                    }
                else{
                        return equationType = "ONLY"                        
                    }
                    }            
            else{
                return equationType = "LAST"
                }            
            }
           
        
    }



    function applyStyletoEquation()
    {
        var my_image = myDoc.allGraphics; 
        for(var i=0;i<my_image.length;i++)
        {
            var mySel = my_image[i].itemLink.name;
            var myEquPath = my_image[i].itemLink.filePath;
            var myID = myEquPath.toString().split ("/").pop().split ("_").pop().replace (/\.eps/gi, "");
            var Regx = /(Equ[0-9]*)/gi;
            if(myID.match(Regx))
            {
                var myEqu = my_image[i];
                myEqu.applyObjectStyle(myDoc.objectStyles.item("EquationFrameObjectStyle"), true)
            }
        }
    }

	function replaceDoubleEnter(myDoc)
	{
		app.findGrepPreferences = NothingEnum.nothing;
		app.changeGrepPreferences = NothingEnum.nothing;
		app.findChangeGrepOptions.includeFootnotes = false;
		app.findChangeGrepOptions.includeHiddenLayers = false;
		app.findChangeGrepOptions.includeLockedLayersForFind = false;
		app.findChangeGrepOptions.includeLockedStoriesForFind = false;
		app.findChangeGrepOptions.includeMasterPages = false;
		app.findGrepPreferences.findWhat = "\r\r";
		app.changeGrepPreferences.changeTo = "";
		myDoc.changeGrep();
	}

    function AddQueryReferences(myDoc)
    {
        logMsg = logMsg + (++msgCounter) + ". Add Query refrences. " + "\n"
        var place = getNodes(myDoc, ".//cs_repos[@typename='Query']");
        for(var i=0;i<place.length;i++)
        {
            var SBnd =place[i];
            var TFooter = false;
            var tableframe  = null;
            var cSBRefNode  = SBnd;
            var boxFigure  = false;
            var RespoidFig = cSBRefNode.xmlAttributes.item("repoID").value;
            var insp = SBnd.insertionPoints[0];
            var pp=null;
            var authorQueryFrame=null;
            try{
                authorQueryFrame=insp.parentTextFrames[0]
                var Page = insp.parentTextFrames[0].parent;
                pp=insp.parentTextFrames[0].parentPage;
            }catch(e){}
            if (insp.parentTextFrames[0].label.toLowerCase()=="tableframe" && cSBRefNode.parent.parent.markupTag.name == "tfooter")
            {
                TFooter = true;
                tableframe = insp.parentTextFrames[0];
                tableframe.geometricBounds = [tableframe.geometricBounds[0], tableframe.geometricBounds[1], tableframe.geometricBounds[2] + 15, tableframe.geometricBounds[3]];
            }
            var AncFrame = insp.textFrames.add();
            if(authorQueryFrame!=null)
                authorQueryFrame.fit(FitOptions.FRAME_TO_CONTENT);
            var currentTFrame=insp.parentTextFrames[0];
            var pageLeftMargin=Math.round(pp.marginPreferences.left);
            var currentTFrameLeft=Math.round(currentTFrame.geometricBounds[1]);
            var AncObjSet = AncFrame.anchoredObjectSettings;
            AncObjSet.anchoredPosition = AnchorPosition.anchored;
            AncObjSet.anchorPoint = AnchorPoint.centerAnchor;
            AncObjSet.horizontalReferencePoint = AnchoredRelativeTo.pageEdge;
            AncObjSet.verticalReferencePoint = VerticallyRelativeTo.lineXheight;
            if(insp.parentTextFrames[0]!=null)
            if (insp.parentTextFrames[0].label.toLowerCase() == "main1") 
            {
                try{
                    AncFrame.applyObjectStyle(myDoc.objectStyles.item("Sidebar_right"), true);
                }catch(e){}
            } 
            else
            {  
                if(currentTFrameLeft==pageLeftMargin)
                {
                    try{
                            AncFrame.applyObjectStyle(myDoc.objectStyles.item("Sidebar_left"), true);
                    }catch(e){}
                }  
                else
                {
                    try{
                        AncFrame.applyObjectStyle(myDoc.objectStyles.item("Sidebar_right"), true);
                    }catch(e){}   
                }   
            }
            if(authorQueryFrame!=null)
                authorQueryFrame.fit(FitOptions.FRAME_TO_CONTENT);
            if (TFooter == true)
                tableframe.geometricBounds = [tableframe.geometricBounds[0], tableframe.geometricBounds[1], tableframe.geometricBounds[2] - 15, tableframe.geometricBounds[3]];
            if (language.toLowerCase() == "en" )
                AncFrame.contents = "AQ" + RespoidFig;
            else
                AncFrame.contents = "FA" + RespoidFig;
            AncFrame.label = "name=AuthorQuery;position=in;Respoid=" + RespoidFig + ";";
            AncFrame.geometricBounds = [AncFrame.geometricBounds[0], AncFrame.geometricBounds[1], AncFrame.geometricBounds[0] +15, AncFrame.geometricBounds[1] + 30];
            AncFrame.fit(FitOptions.FRAME_TO_CONTENT);
        }
        DisplayAuthorQuery(myDoc);
    }

    function BiographyImagePlacement()
    {
        try{
            var myScriptPath = File (sAppPath+"/"+"Scripts/Breeze/BiographyImagePlacementJournals.jsxbin");
            app.doScript(myScriptPath);
        }catch(e){}
    }

    function DisplayAuthorQuery(myDoc)
    {
        var place = getNodes(myDoc, ".//AQChapter");
        var authNode;
        var articleId= FolderName.toString().split("_")[2];
        var autherQueryVariables = myDoc.textVariables;
        autherQueryVariables.item("Journal").variableOptions.contents = myTemplateJournalID.toString();
        autherQueryVariables.item("Article").variableOptions.contents = articleId;
        for(var i=0;i<place.length;i++)
        {
            authNode=place[i];
            if (authNode.contents!="")
            {
                var addBlankPage = myDoc.pages.add(LocationOptions.AT_END,  page);
                if(TemplateName=="3_Col_Journals.indt" ||TemplateName=="2_Col_Journals.indt" || TemplateName=="1_Col_Journals.indt")
                {
                    var masterPage=null;
                    if(language.toLowerCase()=="en")
                        masterPage=getMasterPage("C-English_AQ");  
                    else if(language.toLowerCase()=="de" || language.toLowerCase()=="nl")  
                        masterPage=getMasterPage("D-German_AQ");  
                    if(masterPage==null && !masterPage.isValid)  
                        masterPage=getMasterPage ("Master")  
                    addBlankPage.appliedMaster=masterPage;
                }        
                else
                    addBlankPage.appliedMaster=myDoc.masterSpreads.item("A-Normal Page"); 
                detachFrameWithlabel ("AuthorQuery", addBlankPage)
                var AQFrame = null;
                var allTextFrames=addBlankPage.textFrames;
                for(var t=0;t<allTextFrames.length;t++)
                {
                    if(allTextFrames[t].label.toLowerCase()=="authorquery")
                    $.writeln ("001:" + allTextFrames[t].contents);
                        AQFrame= allTextFrames[t];
                }
                
                if(AQFrame!=null)
                {
                    AQFrame.placeXML(authNode);
                    myDoc.save(SavePath); 
                    while(AQFrame.overflows)
                    {
                        
                        var newBlankPage = myDoc.pages.add(LocationOptions.AT_END,  addBlankPage);
                        var masterPage=null;
                        if(language.toLowerCase()=="en")
                            masterPage=getMasterPage("F-English_AQ");  
                        else if(language.toLowerCase()=="de")  
                            masterPage=getMasterPage("G-German_AQ");  
                        if(masterPage==null && !masterPage.isValid)  
                            masterPage=getMasterPage ("Master")  
                        newBlankPage.appliedMaster=masterPage;
                        detachFrameWithlabel ("AuthorQuery", newBlankPage)
                        var nextAQFrame = null;
                        var allTextFrames=newBlankPage.textFrames;
                        for(var t=0;t<allTextFrames.length;t++)
                        {
                            if(allTextFrames[t].label.toLowerCase()=="authorquery")
                                nextAQFrame = allTextFrames[t];
                        }
                        AQFrame.nextTextFrame=nextAQFrame;
                        AQFrame=nextAQFrame;
                    }
                }
            }
        }
    }

    function moveAutherQueryFrame(authorQueryFrame)
    {
        var authorQueryPage=authorQueryFrame.parentPage
        var xCor,yCor=0;
        switch(authorQueryPage.side)
        {
            case PageSideOptions.LEFT_HAND: yCor=authorQueryPage.marginPreferences.top, xCor =(myDoc.documentPreferences.pageWidth*-1);
                break;
            case PageSideOptions.RIGHT_HAND: yCor=authorQueryPage.marginPreferences.top, xCor =(myDoc.documentPreferences.pageWidth+authorQueryPage.marginPreferences.right);
                break;
        }
        authorQueryFrame.move(authorQueryPage);
        authorQueryFrame.move([xCor,yCor]);     
     }
 

    function Deletepdf()
    {
        logMsg = logMsg + (++msgCounter) + ". Delete existing PDF. " + "\n"
        mypdfFolder = new Folder(PdfPath)
        myExitpdf = mypdfFolder.getFiles ("*");
        for(var p=0; myExitpdf.length>p; p++)
        {
            myExitpdf[p].remove();
        }
    }

    function InlineEquation(myDoc, EqNode)
    {
        var Insp = EqNode.insertionPoints[0];
        var EqPath  = EqNode.xmlAttributes.item("cs_EqPath").value;
        var rect = null;
        if (EqPath != "") 
        {
            rect = Insp.place(new File(EqPath));
            rect [0].parent.appliedObjectStyle = myDoc.objectStyles.item("InlineEqFrameObjectStyle");
        }    
        else
        {          
            rect = Insp.place(new File("c:\EquationMissing.eps"));
            rect [0].parent.appliedObjectStyle = myDoc.objectStyles.item("InlineEqFrameObjectStyle");
        }   
        rect [0].parent.appliedObjectStyle = myDoc.objectStyles.item("InlineEqFrameObjectStyle");
        if(EqPath != ""){
            rect[0].itemLink.unlink()
        }

        var eqNodeParent = EqNode.parent;
        ensureParagraphStyle(eqNodeParent,myDoc,null);
        myDoc.save(SavePath);
    }

    function ensureParagraphStyle(indXmlElement, doc, txtf) 
    {
        if (indXmlElement != null && indXmlElement.xmlAttributes != null) 
        {
            var aidPstyle = indXmlElement.xmlAttributes.item("aid:pstyle");
            var val = "";
            if (aidPstyle != null) 
            {
                if(aidPstyle.value.constructor.name=="Array")
                {
                    val = aidPstyle.value[0];
                    for ( var i = 0; i <indXmlElement.paragraphs.length; i++) 
                    {
                        var docvalue= doc.paragraphStyles.item( val);
                        var apstyle=indXmlElement.paragraphs[i];
                        try{
                            if(apstyle!=null && apstyle.appliedParagraphStyle!=null )
                                apstyle.appliedParagraphStyle =docvalue ;
                        }catch(e){ 
                            alert(e);
                        }
                    }
                }
                else
                {
                    val = aidPstyle.value;
                    if(val=="")
                    for ( var i = 0; i <indXmlElement.paragraphs.length; i++) 
                    {
                        var docvalue= doc.paragraphStyles.item( val);
                        indXmlElement.paragraphs[i].appliedParagraphStyle =docvalue ;
                    }
                }
            }
        }
    }

    function AddFrame2FigHyper( myDoc)
    {
        logMsg = logMsg + (++msgCounter) + ". Add hyperlink frame to figure. " + "\n"
        myDoc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.millimeters;
        myDoc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.millimeters;
        var PageCount  = myDoc.pages.length;
        for(var i  = 0;i<PageCount;i++)
        {
            var page= myDoc.pages[i];
            for (var j  = 0;j< page.groups.length;j++)
            { 
                var gp =page.groups[j];
                var groupLabel=gp.label;
                var HyperTxtf = page.textFrames.add();
                HyperTxtf.geometricBounds = [gp.geometricBounds[0]+4 , gp.geometricBounds[1] + 4  , gp.geometricBounds[0], gp.geometricBounds[1]];
                HyperTxtf.contents = " ";
                HyperTxtf.textFramePreferences.ignoreWrap = true;          
                HyperTxtf.label = groupLabel;
            }
        }
        makeGroupWithHyperlink();
    }

    function DisplayHyperlink(inddPdfDoc)
    {
        var ReqNds= getNodes(inddPdfDoc,".//*[@cs_srcRef]");
        for(var i  = 0;i< ReqNds.length;i++)
        {
            var srcNode  = ReqNds[i];
            var srcAttr  = srcNode.xmlAttributes.item("cs_srcRef");
            var srcAttrValue = srcNode.xmlAttributes.item("cs_srcRef").value;
            var destNds = getNodes(inddPdfDoc,".//*[@cs_destRef='" + srcAttrValue  + "']");
            if (destNds != null) 
            {
                for (var j  = 0;j< destNds .length;j++)
                {
                    var count  = destNds .length - j;
                    var desNode = destNds [j];
                    var destAttr = desNode.xmlAttributes.item("cs_destRef");
                    var allDestNodes=desNode.xmlElements;
                    if(desNode.parent.markupTag.name=="cs_CaptionWrapper")
                    {
                        for(var k=0;k<allDestNodes.length;k++)
                        {
                            if(allDestNodes[k].markupTag.name=="CaptionNumber" )
                            {
                                var capNumberNode=allDestNodes[k];
                                addHyper_Page(inddPdfDoc, srcNode,desNode, count,capNumberNode);
                            }
                        }
                    }  
                }
            }
        }
        for(var j  = 0;j<inddPdfDoc.hyperlinks.length;j++)
        {
            inddPdfDoc.hyperlinks[j].visible = false;
        }
    }
    
    function DisplayHyperlinkForCitation(inddPdfDoc)
    {
        logMsg = logMsg + (++msgCounter) + ". Display Hyperlinks. " + "\n"
        var ReqNds= getNodes(inddPdfDoc,".//*[@cs_srcRef]");
        for(var i  = 0;i< ReqNds.length;i++)
        {
            var srcNode  = ReqNds[i];
            var srcAttr  = srcNode.xmlAttributes.item("cs_srcRef");
            var srcAttrValue = srcNode.xmlAttributes.item("cs_srcRef").value;
            var destNds = getNodes(inddPdfDoc,".//*[@cs_destRef='" + srcAttrValue  + "']");
            if (destNds != null) 
            {
                for (var j  = 0;j< destNds .length;j++)
                {
                    var count  = destNds .length - j;
                    var desNode = destNds [j];
                    var destAttr = desNode.xmlAttributes.item("cs_destRef");
                    var allDestNodes=desNode.xmlElements;
                    if(desNode.parent.markupTag.name=="cs_CaptionWrapper")
                    {
                        for(var k=0;k<allDestNodes.length;k++)
                        {
                            if(allDestNodes[k].markupTag.name=="CaptionNumber" )
                                continue;
                        }
                    }  
                    else
                        addHyper(inddPdfDoc, srcNode, desNode, count);
                }
            }
        }
        for(var j  = 0;j<inddPdfDoc.hyperlinks.length;j++)
        {
            inddPdfDoc.hyperlinks[j].visible = false;
        }
    }

    function addHyper_Page(myDoc, src, desti, count,capNumberNode )
    {
        var hss = myDoc.hyperlinkTextSources;
        var hdd = myDoc.hyperlinkTextDestinations;
        var indTS = src.texts[0];
        var hs = hss.add(indTS);
        var indTD =null;
        if(desti.insertionPoints[0].parentTextFrames[0]!=null && desti.insertionPoints[0].parentTextFrames[0]!=null && desti.insertionPoints[0].parentTextFrames[0].parent!=null)
        {	
            var FigureGroupName  = desti.insertionPoints[0].parentTextFrames[0].parent.label;
            FigureGroupName = FigureGroupName.replace("ImageGroup", "ImageHyperRef");
        }
        if(desti.insertionPoints[0].parentTextFrames[0]!=null &&desti.insertionPoints[0].parentTextFrames[0].parent.parent!=null)
        {
            var Parentpage = desti.insertionPoints[0].parentTextFrames[0].parentPage;
            for(var i=0;i<Parentpage.textFrames.length;i++)
            {
                var tf= Parentpage.textFrames[i];       
                if (tf.label == capNumberNode.contents) 
                {
                    indTD = tf.texts[0];		
                    break;
                }
            }
            if (tf.label == "table") 
            {
                if(tf.characters[count]!=null)
                    indTD = tf.characters[count].texts[0];			
            }
        }
        if((indTD!= null))
        {
            var  hd = hdd.add(indTD);
            myDoc.hyperlinks.add(hs, hd);
        }
    }

    function addHyper(myDoc, src , desti ,  count )
    {
        var hss = myDoc.hyperlinkTextSources;
        var hdd  = myDoc.hyperlinkTextDestinations;
        var indTS = src.texts[0];
        var hs = hss.add(indTS);
        var indTD = desti.characters[count].texts[0];
        var hd = hdd.add(indTD);
        myDoc.hyperlinks.add(hs, hd);
    }

    function makeGroupWithHyperlink()
    {
        var pages=myDoc.pages;
        for(var i=0;i<pages.length;i++)
        {
            var groups=pages[i].groups;
            for(var j=0;j<groups.length;j++)
            {
                var textFrames=pages[i].textFrames;
                for(var k=0;k<textFrames.length;k++)
                { 
                    var label=textFrames[k].label;
                    if(groups[j].label!=""&&groups[j].label!="TipEyecatcherBox"&&groups[j].label!="Definition")
                    {
                        if(textFrames[k].label==groups[j].label)
                        {
                            var imageGroup=pages[i].groups.add([groups[j],textFrames[k]]);
                            imageGroup.label=label;
                        }
                    }
                }
            }
        }    
    }

    function removeBlankBoxFrame()
    {
        var docHeight=myDoc.documentPreferences.pageHeight;                 
        var allPages=myDoc.pages;
        for(var i=allPages.length-1;i>=0;i--)
        {
            var page=allPages[i];
            var bottom=page.marginPreferences.bottom;
            var allTextFrames=page.textFrames;
            for(var j=0;j<allTextFrames.length;j++)
            {
                if(allTextFrames[j].label.indexOf("AnchorFrame")!=-1)
                {
                    if(allTextFrames[j].contents=="")
                        allTextFrames[j].remove ();
                    if(allTextFrames[j].contents!="")
                    {
                        var oldFrameWidth=allTextFrames[j].geometricBounds[3];
                        allTextFrames[j].fit(FitOptions.frameToContent);
                        var newFrameWidth=allTextFrames[j].geometricBounds[3];
                        if(oldFrameWidth>=newFrameWidth)
                            allTextFrames[j].geometricBounds=[allTextFrames[j].geometricBounds[0],allTextFrames[j].geometricBounds[1],allTextFrames[j].geometricBounds[2],oldFrameWidth];
                        var textFrameHeight=allTextFrames[j].geometricBounds[2]-allTextFrames[j].geometricBounds[0];
                        if(docHeight-bottom+2<textFrameHeight)
                            allTextFrames[j].geometricBounds=[allTextFrames[j].geometricBounds[0],allTextFrames[j].geometricBounds[1],docHeight-bottom,allTextFrames[j].geometricBounds[3]];
                        var textF=allTextFrames[j];
                        while(textF.overflows==true)
                        {
                            var txt=findPlacedTextFrame (textF);
                            var newTextFrame= txt.nextTextFrame.parentPage.textFrames.add(); 
                            if(txt.parentPage.name!=txt.nextTextFrame.parentPage.name)
                            {
                                if(parseInt(txt.nextTextFrame.parentPage.name)%2==0)
                                    newTextFrame.geometricBounds=[textF.geometricBounds[0],page.marginPreferences.right,textF.geometricBounds[2],textF.geometricBounds[1]-page.marginPreferences.columnGutter+page.marginPreferences.right-page.marginPreferences.left]
                                if(parseInt(txt.nextTextFrame.parentPage.name)%2!=0)
                                    newTextFrame.geometricBounds=[textF.geometricBounds[0],page.marginPreferences.left,textF.geometricBounds[2],textF.geometricBounds[1]-page.marginPreferences.columnGutter+page.marginPreferences.left-page.marginPreferences.right]
                            }
                            else
                            {
                                if(parseInt(txt.nextTextFrame.parentPage.name)%2==0)
                                    newTextFrame.geometricBounds=[textF.geometricBounds[0],newTextFrame.geometricBounds[3]+page.marginPreferences.columnGutter,textF.geometricBounds[2],textF.geometricBounds[3]]
                                if(parseInt(txt.nextTextFrame.parentPage.name)%2!=0)
                                    newTextFrame.geometricBounds=[textF.geometricBounds[0],newTextFrame.geometricBounds[3]+page.marginPreferences.columnGutter,textF.geometricBounds[2],textF.geometricBounds[3]]
                            }
                            newTextFrame.appliedObjectStyle = myDoc.objectStyles.item("ProcedureBK");
                            textF.nextTextFrame=newTextFrame;
                            var oldFrameWidth=newTextFrame.geometricBounds[3];
                            newTextFrame.fit(FitOptions.frameToContent);
                            var newFrameWidth=newTextFrame.geometricBounds[3];
                            if(oldFrameWidth>=newFrameWidth)
                                newTextFrame.geometricBounds=[newTextFrame.geometricBounds[0],newTextFrame.geometricBounds[1],newTextFrame.geometricBounds[2],oldFrameWidth];
                        }
                    }
                }
            }
        }
    }

    function roman2arabic (roman)
    {
        var arabic = rom2arab (roman.substr (-1));
        for (var i = roman.length-2; i > -1; i--)
        if (rom2arab (roman[i]) < rom2arab (roman[i+1]))
            arabic -= rom2arab (roman[i]);
        else
            arabic += rom2arab (roman[i]);
        return arabic
    }

    function rom2arab (rom_digit)
    {
        switch (rom_digit)
        {
            case "i": return 1;
            case "v": return 5;
            case "x": return 10;
            case "l": return 50;
            case "c": return 100;
            case "d": return 500;
            case "m": return 1000;
            case "I": return 1;
            case  "V": return 5;
            case  "X": return 10;
            case  "L": return 50;
            case  "C": return 100;
            case  "D": return 500;
            case "M": return 1000;  
            default: return "Illegal character"
        }
    }
                  
    function Proofpdf(myDoc)
    {
        logMsg = logMsg + (++msgCounter) + ". Create Proof PDF. " + "\n"
        var PDFSave = PdfPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        var myFileName = myDoc.name.replace(/\.indd$/i, "");
        var Print_Presets = app.pdfExportPresets;
        var ExPDFs = File(PDFSave +"/"+myFileName + "_Proof.pdf");
        var AQExPDFs = File(PDFSave +"/"+myFileName + "_AQ.pdf");
        var FnlPdfPath = ExPDFs.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        var AQFnlPdfPath = AQExPDFs.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        for(var i=0; i<Print_Presets.length; i++)
        {
            if((Print_Presets[i].name == "SpringerOnline_1003_Acro8_Hyperlink_Bookmark"))
              pdfPreset = app.pdfExportPresets.item(Print_Presets[i].name);
        }                     
        if ((pdfPreset.name == "SpringerOnline_1003_Acro8_Hyperlink_Bookmark"))
        {
            var docLayers = myDoc.layers;
            for (var i = 0; i < docLayers.length; i++) 
            {
                var layername = myDoc.layers[i].name
                if (myDoc.layers[i].name=="WaterMark")
                    myDoc.layers[i].visible=true;              
            }
            try{
                var authorQueryNode=getNodesFromAttributeTypeAndValue("AQChapter", "", "", myDoc.xmlElements[0]);
                var  proofPageRange=null;
                var  aqPageRange=null;
                if(authorQueryNode!=null)
                {
                    var aqStartPage=findPage (authorQueryNode.insertionPoints[0]).name;                        
                    var aqEndPage=findPage (authorQueryNode.insertionPoints[-1]).name;                        
                    proofPageRange="1-"+(aqStartPage-1);
                    aqPageRange=aqStartPage+"-"+aqEndPage;
                }
                var presetName=pdfPreset.name; 
                if(aqPageRange!=null)
                {
                    app.pdfExportPreferences.pageRange = proofPageRange;
                    app.documents.item(0).exportFile(ExportFormat.pdfType, FnlPdfPath, app.pdfExportPresets.item(presetName.toString()), false);
                    app.pdfExportPreferences.pageRange = aqPageRange;
                    app.documents.item(0).exportFile(ExportFormat.pdfType, AQFnlPdfPath, app.pdfExportPresets.item(presetName.toString()), false);
                }
                else
                {
                    app.pdfExportPreferences.pageRange = PageRange.ALL_PAGES;
                    app.documents.item(0).exportFile(ExportFormat.pdfType, FnlPdfPath, app.pdfExportPresets.item(presetName.toString()), false);
                }
                
            }catch(e){
                if(aqPageRange!=null)
                {
                    app.pdfExportPreferences.pageRange = proofPageRange;
                    myDoc.exportFile(ExportFormat.pdfType, FnlPdfPath, false, pdfPreset); 
                    app.pdfExportPreferences.pageRange = aqPageRange;
                    myDoc.exportFile(ExportFormat.pdfType, AQFnlPdfPath, false, pdfPreset); 
                }
                else
                {
                    app.pdfExportPreferences.pageRange = PageRange.ALL_PAGES;
                    myDoc.exportFile(ExportFormat.pdfType, FnlPdfPath, false, pdfPreset); 
                }
            }
            app.pdfExportPreferences.pageRange = "";
            for (var i = 0; i < docLayers.length; i++) 
            {
                var layername = myDoc.layers[i].name
                if (myDoc.layers[i].name=="WaterMark")
                    myDoc.layers[i].visible=false;              
            }
            _gEndPageNum = app.documents.item(0).pages.length;
        } 
    }

    function Hyperlink_CitationRef(myDoc)
    {
        logMsg = logMsg + (++msgCounter) + ". Hyperlink Citiation refrences. " + "\n"
        var ReqNds=getNodes(myDoc, ".//CitationRef");
        for(var  i = 0;i< ReqNds.length;i++)
        {
            node = ReqNds[i];
            var ChaLength=ReqNds[i].contents;
            var CitationRef_Contents = ReqNds[i];
            for(var len =0; len<ChaLength.length; len++)
            {
                try{
                    CitationRef_Contents.characters[len].appliedCharacterStyle = myDoc.characterStyles.item("hyperlink");
                }catch(e){}
            }    
        }  
    }

    function Hyperlink_InternalRef(myDoc)
    {
        logMsg = logMsg + (++msgCounter) + ". Hyperlink internal Refrence. " + "\n"
        var ReqNds= getNodes(myDoc,".//InternalRef");
        var TabRex=/(Tab[1-9]*)/gi;
        var FigRex=/(Fig[1-9]*)/gi;
        for(var i=0;i<ReqNds.length;i++)
        {
            var node=ReqNds[i];
            var nodeAttrVariable=node.xmlAttributes.item("RefID").value;
            if(nodeAttrVariable.match(TabRex) || nodeAttrVariable.match(FigRex))
            {
                var ChaLength=node.contents;
                var InternalRef_Contents = ReqNds[i];
                for(var len =0; len<ChaLength.length; len++)
                {
                    try{
                        InternalRef_Contents.characters[len].appliedCharacterStyle = myDoc.characterStyles.item("Callout-hyperlink");
                    }catch(e){}
                }            
            } 
        }
    }

    function ChangeColor(myDoc)
    {
        try{
            app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
            app.findChangeGrepOptions.includeFootnotes=true;
            app.findChangeGrepOptions.includeHiddenLayers=true;
            app.findChangeGrepOptions.includeLockedLayersForFind=true;
            app.findChangeGrepOptions.includeMasterPages=true;
            app.findGrepPreferences.fillColor = "Hyperlink";
            app.changeGrepPreferences.fillColor = "Black";
            var changeResult = myDoc.changeGrep();
            app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
        }catch(e){}
        myDoc.save();
    }
    
    function ZipPdf()
    {
        var MoveFile = new File("D:/Breeze"+"/"+"ZipPdf.bat" );
        MoveFile.open("w");
        MoveFile.write("D:"+"\n"+"cd Breeze"+"\n"+"java -jar ZipPdf1.0.jar "+FolderName+"\n"+" exit");
        MoveFile.close();
        MoveFile.execute(); 
        $.sleep (2000)
    } 

    function MergePdf()
    {
        var MoveFile = new File("D:/Breeze"+"/"+"MergePdf.bat" );
        MoveFile.open("w");
        MoveFile.write("D:"+"\n"+"cd Breeze"+"\n"+"java -jar MergePdf1.0.jar"+" "+'"'+"D:\\Breeze-Pagination\\SPR_Mediz_T5a_M_2Col\\IN\\"+FolderName);
        MoveFile.close();
        MoveFile.execute(); 
        $.sleep (2000)
    } 

    function MoveVoucherAndMetaDataPdf()
    {
        ZipPath = zipPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");   
        myFolder = zipPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        myPath = "//Zapprocess/GeneralInfo/HardDrivePath"
        var myZipPath=File(ZipPath+"/");
        myPGXML = myZipPath.getFiles("*.xml");
        for (var pg = 0; myPGXML.length>pg; pg++)
        {
            if(File(myPGXML[pg].exists))
            {
                temp1 = myPGXML[pg].toString().split("/");
                temp = temp1.pop();
                pgTemp = temp.toString().split("_")
                pgRemp1 = pgTemp.pop();
                var Regx=/(PG.xml)/gi;
                if(temp.match(Regx))
                { 
                    var Docdatafile=File(myPGXML[pg]);
                    Docdatafile.open('r');
                    var fileString = Docdatafile.read();
                    var myPgString = new XML( fileString );
                    var myString=myPgString.xpath(myPath);
                    Docdatafile.close();
                    var MovePdfFiles=new Folder(myString);
                    var voucherMetaDataPdfPath=ZipPath+"/"+"PDF"+"/";
                    var myVoucherMetaDataPdfPath=File(voucherMetaDataPdfPath+"/");
                    var myVoucherMetaPdf=myVoucherMetaDataPdfPath.getFiles ("*.pdf");
                    for(var i=0;i<myVoucherMetaPdf.length;i++)
                    {
                        if(File(myVoucherMetaPdf[i].exists))
                        {
                            temp1 = myVoucherMetaPdf[i].toString().split("/");
                            temp = temp1.pop();
                            pgTemp = temp.toString().split("_")
                            pgRemp1 = pgTemp.pop();
                            var Regx_Proof=/(ProofVoucher.pdf)/gi;
                            var Regx_Metadata=/(XmlMetaData.pdf)/gi;
                            var  copyPdfPath=myString+"\\"+"Attachments";
                            var myStringFolder=new Folder(copyPdfPath);
                            if (!myStringFolder.exists)
                                myStringFolder.create();
                            if(temp.match(Regx_Proof) || temp.match(Regx_Metadata))
                                var myResult=myVoucherMetaPdf[i].copy(myStringFolder +"\\" +myVoucherMetaPdf[i].name);
                        }   
                    }
                    var paginationPath=ZipPath+"/"+"Pagination"+"/";
                    var myPaginationPath=File(paginationPath+"/");  
                    var indesignDocumnetFile=myPaginationPath.getFiles ("*.indd");
                    for(var i=0;i<indesignDocumnetFile.length;i++)
                    {
                        if(File(indesignDocumnetFile[i].exists))
                        {
                            myString=myString.split("Chapters")[0]+"Chapters"+myString.split("Chapters")[1].substring (0, 2);
                            var  paginationCopyPath=myString+"\\"+"Pagination";
                            var myStringFolder=new Folder(paginationCopyPath);
                            if (!myStringFolder.exists)
                                myStringFolder.create();
                           var myResult=indesignDocumnetFile[i].copy(myStringFolder+"\\"+indesignDocumnetFile[i].name.split(".")[0]+"_Autopaginated.indd");
                        }   
                    }
                }
            }
        }
    }

    function MoveZip()
    {
        ZipPath = zipPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        myFolder = zipPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        myPath = "//Zapprocess/GeneralInfo/UserHomeDrivePath"
        var myZipPath=File(ZipPath+"/");
        myPGXML = myZipPath.getFiles("*.xml");
        for (var pg = 0; myPGXML.length>pg; pg++)
        {
            if(File(myPGXML[pg].exists))
            {
                temp1 = myPGXML[pg].toString().split("/");
                temp = temp1.pop();
                pgTemp = temp.toString().split("_")
                pgRemp1 = pgTemp.pop();
                var Regx=/(PG.xml)/gi;
                if(temp.match(Regx))
                {
                    var Docdatafile=File(myPGXML[pg]);
                    Docdatafile.open('r');
                    var fileString = Docdatafile.read();
                    var myPgString = new XML( fileString );
                    var myString=myPgString.xpath(myPath);
                    Docdatafile.close();
                    var MoveZipFolder=new Folder(myString);
                    var myinddfolder = new Folder(myFolder);
                    if (!MoveZipFolder.exists)
                        MoveZipFolder.create();
                    var myFiles = myinddfolder.getFiles( "*.zip" );
                    for ( i = myFiles.length-1; i >= 0 ; i-- )
                    {
                        var myResult = myFiles[i].copy(MoveZipFolder +"/" + myFiles[i].name);
                    }
                }
            }
        }
    }

    function MovePdflog()
    {
        ZipPath = zipPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        myFolder = zipPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        myPath = "//Zapprocess/GeneralInfo/UserHomeDrivePath"
        var myZipPath=File(ZipPath+"/");
        myPGXML = myZipPath.getFiles("*.xml");
        for (var pg = 0; myPGXML.length>pg; pg++)
        {
            if(File(myPGXML[pg].exists))
            {
                temp1 = myPGXML[pg].toString().split("/");
                temp = temp1.pop();
                pgTemp = temp.toString().split("_")
                pgRemp1 = pgTemp.pop();
                var Regx=/(PG.xml)/gi;
                if(temp.match(Regx))
                {
                    var Docdatafile=File(myPGXML[pg]);
                    Docdatafile.open('r');
                    var fileString = Docdatafile.read();
                    var myPgString = new XML( fileString );
                    var myString=myPgString.xpath(myPath);
                    Docdatafile.close();
                    var MoveZipFolder=new Folder(myString);
                    var myinddfolder = new Folder(myFolder);
                    if (!MoveZipFolder.exists)
                        MoveZipFolder.create();
                    var myFiles = myinddfolder.getFiles( "*_Error.txt" );
                    for ( i = myFiles.length-1; i >= 0 ; i-- )
                    {
                        var myResult = myFiles[i].copy(MoveZipFolder +"/" + myFiles[i].name);
                    }
                }
            }
        }
    }

    function  createLog()
    {                 
        myfilePath= new File(basePath+"/");
        mylogpath = myfilePath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        var successFile= new File(mylogpath +"/" + myLogFileName+"_Success.txt")
        if(successFile.open("w", undefined, undefined))
        {
            successFile.write((_gStartPageNum + _gAddCount)+"-"+(_gAddCount + _gEndPageNum));
        }
        successFile.close();
    }

    function PgXml()
    {
        myOutPathtemp= new File("D:/Breeze-Pagination/SPR_Mediz_T5a_M_2Col/OUT/");
        myOutPath = myOutPathtemp.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        var myMoveFile=File("D:/Breeze" +"/"+ FolderName+ "_MoveFolder.txt");
        myPath = "//Zapprocess/GeneralInfo/UserHomeDrivePath"
        var basePath = FolderList.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        myMoveFile.open("w");
        myMoveFile.write(basePath+"\n"+myOutPath+"/"+FolderName);
        myMoveFile.close();
    }

    function MoveFolder()
    {
        var myMoveTxt=File(FolderPath +"/"+ FolderName+ "_MoveFolder.txt");
        TxtPath = myMoveTxt.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        var MoveFile = new File("D:/Breeze"+"/"+FolderName +"_MoveFolder.bat" );
        MoveFile.open("w");
        MoveFile.write("D:"+"\n"+"cd Breeze"+"\n"+"java -jar MoveFolder1.0.jar"+" " +"D:/Breeze"+"/" + FolderName + "_MoveFolder.txt"+ "\n\n\n\n" +"del" +" " +"D:/Breeze"+"/" + FolderName + "_MoveFolder.txt" + "\n\n\n\n" +"exit");
        MoveFile.close();
        MoveFile.execute();
        $.sleep (200);
        MoveFile.remove();
    }

    function checkLanguage()
    {
        logMsg = logMsg + (++msgCounter) + ". Check Language. " + "\n"
        if(basePath.indexOf("_De_")!=-1)
            language="De";
        if(basePath.indexOf("_En_")!=-1)
            language="En";
        if(basePath.indexOf("_Article")!=-1)
        {
            var articleInfo=getNodes (myDoc, "ArticleInfo")
            if(articleInfo!="")
            {
                var articleInfoAttributes=articleInfo[0].xmlAttributes;
                var languageAttribute=articleInfoAttributes.itemByName("Language");
                if(languageAttribute.isValid)
                {
                    language=languageAttribute.value;
                    chapterLanguage=languageAttribute.value;
                }    
            }
        }        
    }

function getAbstract()
    {
        logMsg = logMsg + (++msgCounter) + ". Get abstract. " + "\n"
        var xPathChapter = "//Chapter"
        var xPathArticle = "//Article"    
        if(Filename.indexOf("FM1")!=-1)
            xPathChapter = "//Book"
        if(Filename.indexOf("Article")!=-1)
            xPathChapter = "//Article"
        var root  = myDoc.xmlElements[0];
        var chapterHeaderNode  = null;
        try 
        {
            var proc  = app.xmlRuleProcessors.add([xPathChapter]);
            var match = proc.startProcessingRuleSet(root);
            while( match!=undefined ||match!=null) 
            { 
                chapterHeaderNode = match.element;
                var chpaterXmlInstructions=chapterHeaderNode.xmlInstructions;
                if(chpaterXmlInstructions.length==0)
                {
                    match = proc.findNextMatch(); 
                     continue;
                }
                for(var i=0;i<chpaterXmlInstructions.length;i++)
                {
                    var abstractData=chpaterXmlInstructions[i];
                    if(abstractData.target=="ADDHDINFO")
                    {
                        abstractNodeData=abstractData.data;
                        break;
                     }
                }
                match = proc.findNextMatch(); 
                break;
            }
        }
        catch( ex ) 
        {
            alert(ex);
        } finally {}
        myAbstractFile=new File(FolderPath +"/"+FolderName.split("_Chapter")[0]+ "_temp.xml");
        myAbstractFile.encoding="UTF-8";   
        var Temps ="<?xml version='1.0' encoding='UTF-8' standalone='no'?>"
        Temps =Temps+"<Book>";
        Temps =Temps+abstractNodeData;
        Temps =Temps+"</Book>";
        myAbstractFile.open("w");
        myAbstractFile.write(Temps)
        myAbstractFile.close();
         myAbstractFile.open('r');
        var temp = myAbstractFile.read();
        myAbstractFile.close();
        temp=temp.replace(/&/gi,'&#x0026;');
         myAbstractFile.open("w");
        myAbstractFile.write(temp);
        myAbstractFile.close();
    }
    
    function globalMetaData(){
        
        var myScriptPath = File (sAppPath+"/"+"Scripts/Breeze/SKAP/AddMetadata.jsxbin");
        myScriptPath = myScriptPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        Fpath=File(myScriptPath);
        myRefresh = Fpath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        app.doScript(myRefresh);
            
        }
    /*function globalMetaData()
    {
        logMsg = logMsg + (++msgCounter) + ". Creating globalmetadata. " + "\n"
        var allFiles=new Folder(basePath).getFiles();
        var chapterFileName=null;
        for(var i=0;i<allFiles.length;i++)
        {
            var xmlMetaDataFile=allFiles[i].fsName; 
            if(xmlMetaDataFile.indexOf("temp")!=-1)
                xmlMetaDataFilePath=allFiles[i].fsName;
        }    
        var myMetaDataDoc =null;
        if(xmlMetaDataFilePath!=null)
        {
            if(chapterLanguage==null)
                chapterLanguage=language;
              if (chapterLanguage.toLowerCase() == "de")
                MetaDatapath = xmlMetaDataTemplatePathDe.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
//~             else if(chapterLanguage.toLowerCase()=="nl")
//~                 MetaDatapath = xmlMetaDataTemplatePathNl.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
             else
                MetaDatapath = xmlMetaDataTemplatePathEn.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
            myMetaDataDoc = app.open(MetaDatapath);
            myMetaDataDoc.viewPreferences.rulerOrigin=RulerOrigin.pageOrigin;   
            myMetaDataDoc.viewPreferences.horizontalMeasurementUnits=MeasurementUnits.MILLIMETERS ; 
            myMetaDataDoc.viewPreferences.verticalMeasurementUnits=MeasurementUnits.MILLIMETERS; 
            
            var myDocVariables = myMetaDataDoc.textVariables
            if(chapterLanguage.toLowerCase() == "de")
            {
                   var JournalFileName = journalPgXmlName.split("_PG")[0];
                   var fileType = JournalFileName.toString().split("_").pop();
                   if(fileType=="Article")
                        myDocVariables.item("Meta Data").variableOptions.contents = "Metadaten des Artikels, die online angezeigt werden";
                  else
                        myDocVariables.item("Meta Data").variableOptions.contents = "Metadaten des Kapitels, die online angezeigt werden";
            }
                   
//~             else if(chapterLanguage.toLowerCase()=="nl")
//~             {
//~                    var JournalFileName = journalPgXmlName.split("_PG")[0];
//~                    var fileType = JournalFileName.toString().split("_").pop();
//~                    if(fileType=="Article")
//~                         myDocVariables.item("Meta Data").variableOptions.contents = "Metadata van het artikel, die online worden weergegeven";
//~                   else
//~                         myDocVariables.item("Meta Data").variableOptions.contents = "Metadata van het hoofdstuk , die online worden weergegeven";
//~             }

            else{
                  var JournalFileName = journalPgXmlName.split("_PG")[0];
                  var fileType = JournalFileName.toString().split("_").pop();
                  if(fileType=="Article")
                        myDocVariables.item("Meta Data").variableOptions.contents = "Metadata of the article that will be visualized online";
                  else
                        myDocVariables.item("Meta Data").variableOptions.contents = "Metadata of the chapter that will be visualized online";
                 }    
        }
       checkMetaDataXmlExist(myMetaDataDoc);

        function checkMetaDataXmlExist(myMetaDataDoc)
        {
            var allFiles=new Folder(basePath).getFiles();
            var chapterFileName=null;
            for(var i=0;i<allFiles.length;i++)
            {
                var xmlMetaDataFile=allFiles[i].fsName; 
                if(xmlMetaDataFile.indexOf("temp")!=-1)
                {
                    xmlMetaDataFilePath=allFiles[i].fsName;
                    var xmlMetaDataFileName=allFiles[i].fsName.split("\\").pop(); 
                    chapterFileName=xmlMetaDataFileName.split("_temp")[0];  
                }
            }
            if(xmlMetaDataFilePath!=null)
                generateChapterMetaData(chapterFlie,chapterFileName,myMetaDataDoc);
        }

        function generateChapterMetaData(chapterFlie,chapterFileName,myMetaDataDoc)
        {
            myMetaDataDoc.importXML(File(xmlMetaDataFilePath));
            createTableForMetaData(myMetaDataDoc)
            refresh();
            xmlMetaDatapdf(myMetaDataDoc,chapterFlie,chapterFileName);
        }

        function createTableForMetaData(myMetaDataDoc)
        {
            var authorNodes=getNodes(myMetaDataDoc,".//M_Author");
            var metaDataNodes=getNodes (myMetaDataDoc, ".//M_BookTitle |.//M_ChapterCopyRightYear|.//M_ArticleTitle |.//M_ArticleCopyRightYear|.//M_ChapterTitle|.//M_Abstract|.//M_Keywords");
            var authorNodesCounts=0;
            if(authorNodes!="")
            {
                    for(var rw = 0; rw < authorNodes.length; rw++)
                    {
                            authorNodesCounts = authorNodesCounts + authorNodes[rw].xmlElements.length;
                    }
            }
            var rowCounter=metaDataNodes.length+14+authorNodesCounts;
            var myTableTag = myMetaDataDoc.xmlTags.add("myTable");
            var myRowTag =myMetaDataDoc.xmlTags.add("myRow");
            var myFirstCellTag = myMetaDataDoc.xmlTags.add("bookTitleCell1");
            var mySecondCellTag = myMetaDataDoc.xmlTags.add("bookTitleCell2");
            var myRootXMLElement = myMetaDataDoc.xmlElements[0]
            var rowInTable=rowCounter;
            if(rowInTable!=0)
            {
                with(myRootXMLElement)
                {
                    var myTableXMLElement = xmlElements.add(myTableTag);
                    with(myTableXMLElement )
                    {
                        for(var myRowCounter = 1;myRowCounter < rowCounter+1;myRowCounter++)
                        {
                            with(xmlElements.add(myRowTag))
                            {
                                for(var myCellCounter = 1; myCellCounter < 5; myCellCounter++)
                                {
                                    with(xmlElements.add(myFirstCellTag))
                                    {
                                        contents =  "" ;
                                    }
                                }
                            }
                        }
                    }
                }
                var myTable = myTableXMLElement.convertElementToTable(myRowTag, myFirstCellTag);
                myTable.bottomBorderStrokeWeight = 0;
                myTable.topBorderStrokeWeight = 0;
                var rows=myTable.rows;
                for(var i=0;i<rows.length;i++)
                {
                    var row=rows[i];
                    row.bottomEdgeStrokeWeight = 0;
                    row.topEdgeStrokeWeight = 0;
                    row.leftEdgeStrokeWeight = 0;
                    row.rightEdgeStrokeWeight = 0;
                    row.innerRowStrokeWeight = 0;
                    row.diagonalLineStrokeWeight = 0;
                    row.innerColumnStrokeWeight = 0;
                    row.topInset = 2;
                    row.bottomInset = 2;
                    row.leftInset = 1;
                    row.rightInset = 2;
                    row.verticalJustification = VerticalJustification.topAlign;
                }
                var tableTextFrame=myTable.parent;
                while(tableTextFrame.overflows==true)
                {
                    var parentPage=tableTextFrame.parentPage;
                    var addBlankPage = myMetaDataDoc.pages.add(LocationOptions.AFTER,  parentPage);
                    addBlankPage.appliedMaster=myMetaDataDoc.masterSpreads.item("A-Master"); 
                    var docWidth=myMetaDataDoc.documentPreferences.pageWidth;
                    var docHeight=myMetaDataDoc.documentPreferences.pageHeight;                 
                    var top=addBlankPage.marginPreferences.top;
                    var bottom=addBlankPage.marginPreferences.bottom;
                    var left=addBlankPage.marginPreferences.left;
                    var right=addBlankPage.marginPreferences.right;
                    var newTextFrame= addBlankPage.textFrames.add();
                    if(parseInt(addBlankPage.name)%2!=0)
                    newTextFrame.geometricBounds=[top,left,docHeight-bottom,docWidth-right];
                    if(parseInt(addBlankPage.name)%2==0)
                    newTextFrame.geometricBounds=[top,right,docHeight-bottom,docWidth-left];    
                    tableTextFrame.nextTextFrame=newTextFrame;
                    tableTextFrame=newTextFrame;
                }
                var metaDataNodes=getNodes (myMetaDataDoc, ".//M_BookTitle | .//M_ChapterCopyRightYear|.//M_ArticleTitle | .//M_ArticleCopyRightYear| .//M_ChapterTitle|.//M_PartTitle");
                for(i=0;i<metaDataNodes.length;i++)
                {
                    var metaDataAttribute=metaDataNodes[i].xmlAttributes.item("type");
                    if(myTable.rows[i].cells[0].contents=="")
                    {
                        if(metaDataAttribute!="")
                            myTable.rows[i].cells[0].contents=metaDataAttribute.value;
                        if(metaDataNodes[i]!="")
                            myTable.rows[i].cells[2].contents=metaDataNodes[i].contents;
                        removeExtraSymbol(myTable.rows[0].cells[2]);
                        myTable.rows[i].cells[0].merge(myTable.rows[i].cells[1]);
                        myTable.rows[i].cells[1].merge(myTable.rows[i].cells[2]);
                        myTable.rows[i].bottomEdgeStrokeWeight = 0.5;
                    }
                }
                var metaDataNodes=getNodes (myMetaDataDoc, ".//M_Author | .//M_Editor|.//M_ProofRecipient/M_ProductionEditor");
                var  k=2;
                var count=0;
                var authorElements=metaDataNodes;
                for(var j=0;j<authorElements.length;j++)
                {
                    count=0;
                    var authorElementsAttribute=authorElements[j].xmlAttributes.item("type");
                    if(count==0)
                    {
                        count=1;
                        myTable.rows[++k].cells[0].contents=authorElementsAttribute.value;
                    }    
                    if(authorElements[j].markupTag.name=="M_Author" ||authorElements[j].markupTag.name=="EditorName" ||authorElements[j].markupTag.name=="ProductionEditorName" )
                    {
                        var authorDetails=authorElements[j].xmlElements;
                        for(var p=0;p<authorDetails.length;p++)
                        {
                            var  authorDetailsAttributeValue=authorDetails[p].xmlAttributes.item("type").value;
                            if(p>0)
                                ++k;
                            if(authorDetailsAttributeValue!="")
                                myTable.rows[k].cells[1].contents=authorDetailsAttributeValue;     
                            var contents=authorDetails[0].insertionPoints[-1].contents;
                            if(authorDetails[p].contents!="")
                                myTable.rows[k].cells[2].contents=authorDetails[p].contents;
                            removeExtraSymbol(myTable.rows[k].cells[2]);
                            myTable.rows[k].cells[2].merge(myTable.rows[k].cells[3]);
                            if(p==authorDetails.length-1)
                                myTable.rows[k].bottomEdgeStrokeWeight = 0.5;
                        }
                    }
                }  
                var abstractText=" ";
                var abstractNode=getNodes (myMetaDataDoc, ".//M_Abstract");
                if(abstractNode!="")
                {
                    abstractPage=abstractNode[0].insertionPoints[0].parentTextFrames[0].parentPage;         
                    var metadataTextFrame=abstractPage.textFrames.add();
                    metadataTextFrame.label="temp";
                    metadataTextFrame.geometricBounds=[0,0,30,100];
                    metadataTextFrame.placeXML(abstractNode[0]);
                    metadataTextFrame.fit(FitOptions.frameToContent);
                    var abstractText=metadataTextFrame.texts[0];
                    var abstractNodeAttribute=abstractNode[0].xmlAttributes.item("type");
                    myAbstractFile.remove();
                    myTable.rows[++k].cells[0].contents=abstractNodeAttribute.value;
                    abstractText.duplicate(LocationOptions.AFTER,myTable.rows[k].cells[2].insertionPoints.lastItem()) ; 
                    metadataTextFrame.remove();
                    myTable.rows[k].cells[2].merge(myTable.rows[k].cells[3]);
                    myTable.rows[k].bottomEdgeStrokeWeight = 0.5;
                }
                var keywordText=""
                var keywordNode=getNodes (myMetaDataDoc, ".//M_Keywords");
                if(keywordNode!="")
                {
                    keywordText=keywordNode[0].contents;
                    var keywordNodeAttribute=keywordNode[0].xmlAttributes.item("type");
                    myTable.rows[++k].cells[0].contents=keywordNodeAttribute.value;
                    myTable.rows[k].cells[2].contents=keywordText;
                    myTable.rows[k].cells[2].merge(myTable.rows[k].cells[3]);
                    myTable.rows[k].bottomEdgeStrokeWeight = 0.5;
                }
                try
                {
                    var removingElement=getNodes(myMetaDataDoc,".//Book")
                    removingElement[0].remove();
                    deleteEmptyRows(myTable);
                }catch(e){
                    alert(e)
                }
                var rows=myTable.rows;
                for(var i=0;i<rows.length;i++)
                {
                    var cells=rows[i].cells;
                    for(var j=0;j<cells.length;j++)
                    {
                        var paragraphs=cells[j].paragraphs;   
                        for(var p=0;p<paragraphs.length;p++)
                        {
                            paragraphs[p].appliedParagraphStyle=myMetaDataDoc.paragraphStyles.item("Author_Detail");      
                        }  
                    }
                }
            }
        }

        function removeExtraSymbol(cell)
        {
            try{
                var s = cell.paragraphs[0];
                var n =  s.length;
                while(s.characters.lastItem().contents==" "||s.characters.lastItem().contents=="\r" )
                {
                    s.characters.lastItem().remove();
                }
            }catch(e){}
        }
        
        function deleteEmptyRows(myTable)
        {
            var tableRows=myTable.rows;
            for(var r=0;r>=0&&r<tableRows.length;r++)
            {
                var deleteFlag=true;
                var singleRowCells=tableRows[r].cells;
                for(var c=0;c<tableRows[r].cells.length;c++)
                {
                    if(!tableRows[r].cells[c].contents=="")
                    {
                        deleteFlag=false;
                        break;
                    }   
                }
                if(deleteFlag)  
                {
                    tableRows[r].remove();
                    r=r-1;
                }   
            }
            if(myTable.isValid)
            {
                var tableFrame=myTable.parent;
                var continueTableFrame=tableFrame.nextTextFrame;
                while(continueTableFrame!=null)
                {
                    if(continueTableFrame.isValid)
                    {
                        continueTableFrame.remove();
                        continueTableFrame=tableFrame.nextTextFrame;
                     }
                     else
                        break;
                }
                deleteEmptyPagesInMetaData(); 
                while(tableFrame.overflows==true)
                {
                    var parentPage=tableFrame.parentPage;
                    var addBlankPage = myMetaDataDoc.pages.add(LocationOptions.AFTER,  parentPage);
                    addBlankPage.appliedMaster=myMetaDataDoc.masterSpreads.item("A-Master"); 
                    var docWidth=myMetaDataDoc.documentPreferences.pageWidth;
                    var docHeight=myMetaDataDoc.documentPreferences.pageHeight;                 
                    var top=addBlankPage.marginPreferences.top;
                    var bottom=addBlankPage.marginPreferences.bottom;
                    var left=addBlankPage.marginPreferences.left;
                    var right=addBlankPage.marginPreferences.right;
                    var newTextFrame= addBlankPage.textFrames.add();
                    if(parseInt(addBlankPage.name)%2!=0)
                        newTextFrame.geometricBounds=[top,left,docHeight-bottom,docWidth-right];
                    if(parseInt(addBlankPage.name)%2==0)
                        newTextFrame.geometricBounds=[top,right,docHeight-bottom,docWidth-left];    
                    tableFrame.nextTextFrame=newTextFrame;
                    tableFrame=newTextFrame;
                }
            }
        }

        function deleteEmptyPagesInMetaData()
        {
            var pages = app.documents[0].pages.everyItem().getElements();
            for(var i = pages.length-1;i>=0;i--)  
            {
                var removePage = true;
                if(pages[i].pageItems.length>0)
                {
                    var items = pages[i].pageItems.everyItem().getElements();
                    for(var j=0;j<items.length;j++)
                    {
                        if((items[j] instanceof TextFrame))
                        {
                            removePage=false;
                            break;
                        }
                   }
                }
                if((removePage&&pages[i].appliedMaster.name=="A-Master")&&app.documents[0].pages.length>0)
                    pages[i].remove()
            }
        }

        function xmlMetaDatapdf(myMetaDataDoc,chapterFlie,chapterFileName)
        {          
            var PdfSave = PdfPath.toString().replace(/\%20/g," ");
            var metaDataPdfSave=PdfSave.replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:") ;
            var Print_Presets = app.pdfExportPresets;
            var ExPDFs =File(metaDataPdfSave+"/"+chapterFileName+ "_XmlMetaData.pdf");
            var ExPDFspath = ExPDFs.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:") ;
            for(var i=0; i<Print_Presets.length; i++)
            {
                if(Print_Presets[i].name == "webpdf"||Print_Presets[i].name == "[webpdf]")
                    pdfPreset = app.pdfExportPresets.item(Print_Presets[i].name);
            }                     
            try{
                var presetName=pdfPreset.name;
                FindChangeByList();
                app.documents.item(0).exportFile(ExportFormat.pdfType, ExPDFspath, app.pdfExportPresets.item(presetName.toString()));
            }
            catch(e){
                myMetaDataDoc.exportFile(ExportFormat.pdfType, ExPDFspath); 
            }  
            _gAddCount = _gAddCount +  app.documents.item(0).pages.length;
            myMetaDataDoc.close(SaveOptions.NO);
            voucherPdf(chapterFileName);
        } 

        function voucherPdf(chapterFileName)
        {
            var proofVoucherDoc=null;
            if(chapterLanguage.toLowerCase()=="de"||chapterLanguage.toLowerCase()=="nl")
            {
                proofVoucherDeTemplatetempPath = proofVoucherDeTemplatePath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
                proofVoucherDoc=app.open(proofVoucherDeTemplatetempPath);
            }
            if(chapterLanguage.toLowerCase()=="en")
                proofVoucherDoc=app.open(proofVoucherEnTemplatePath);
            var PdfSave = PdfPath.toString().replace(/\%20/g," ");
            var proofVoucherPdfSave=PdfSave.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:") ;
            var Print_Presets = app.pdfExportPresets;
            var ExPDFs =File(proofVoucherPdfSave+"/"+chapterFileName+ "_ProofVoucher.pdf");
            for(var i=0; i<Print_Presets.length; i++)
            {
                if(Print_Presets[i].name == "webpdf"||Print_Presets[i].name == "[webpdf]")
                    pdfPreset = app.pdfExportPresets.item(Print_Presets[i].name);
            }                     
            try{
                var presetName=pdfPreset.name;
                app.documents.item(0).exportFile(ExportFormat.pdfType, ExPDFs, app.pdfExportPresets.item(presetName.toString()));
            }catch(e){
                proofVoucherDoc.exportFile(ExportFormat.pdfType, ExPDFs, false, pdfPreset); 
            }   
            _gAddCount = _gAddCount +  app.documents.item(0).pages.length;
            proofVoucherDoc.close(SaveOptions.NO);
        } 
    }*/

    function log(myPGXML, message) 
    {         
        aFile = File(myPGXML);
        var today = new Date();  
        if (!aFile.exists)
        {
            ErrorLog = zipPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
            NewErrorlog = ErrorLog.toString().split("/");
            ErrorLogTemp = NewErrorlog.pop();
            ErrorLogFileName = ErrorLogTemp+"_Error.log";
            aFile = File(ErrorLog+"/"+ErrorLogFileName);
        }
        aFile.open("w"); 
        aFile.write("   Crest PreMedia Solutions Pvt. Ltd.        "+"\n");
        aFile.write("   Dept : Technology (Java Team)        "+"\n");
        aFile.write("***********************************************"+"\n\n");
        aFile.write(" Kindly Check Below Points:-        "+"\n");
        aFile.write("	 1 PGXml Missing        "+"\n");
        aFile.write("	 2 Please Check Mandatory Fields        "+"\n");
        aFile.close();  
    }

    function pdfLog(myPGXML,message) 
    {         
        aFile = File(myPGXML+"/"+ FolderName+"_Error.txt");
        var today = new Date();  
        if (!aFile.exists)
        {
            ErrorLog = zipPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
            NewErrorlog = ErrorLog.toString().split("/");
            ErrorLogTemp = NewErrorlog.pop();
            ErrorLogFileName = FolderName+"_Error.txt";
            aFile = File(ErrorLog+"/"+ErrorLogFileName);
        }
        aFile.open("w"); 
        aFile.write("   Crest PreMedia Solutions Pvt. Ltd.        "+"\n");
        aFile.write("   Dept : Technology (Java Team)        "+"\n");
        aFile.write("***********************************************"+"\n\n");
        aFile.write(" Kindly Check Below Points:-        "+"\n");
        aFile.write("1   "+message+"\n");
        aFile.close();  
    }

    function createFootNotes()
    {
        logMsg = logMsg + (++msgCounter) + ". Creating Footnotes. " + "\n"
        var myScriptPath = File (sAppPath+"/"+"Scripts/Breeze/Footnote_Journal.jsxbin");
        myScriptPath = myScriptPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        app.doScript(myScriptPath);
    }

    function myGetScriptPath()
    {
        try{
            return app.activeScript;
        }catch(myError)
        {
            return File(myError.fileName);
        }
    }
        
    function placeSingleTextWidthImage(imageGroup)
    {
        var isOutSidePageMargin=checkPlacedOutsidePageMargin (imageGroup);
        var imageGroupPage=imageGroup.parentPage;
        if(isOutSidePageMargin)
        {
            if(parseInt(imageGroupPage.name)%2==0)
            {
                var imageGroupPlacedTextFrame=findPlacedTextFrame (imageGroup);
                var nextTextToImageGroupFrame=imageGroupPlacedTextFrame.nextTextFrame;
                var image=imageGroup.rectangles[0];
                var captionFrame=imageGroup.textFrames[0];
                captionFrame.move([captionFrame.geometricBounds[1],imageGroup.geometricBounds[0]]);
                image.move([nextTextToImageGroupFrame.geometricBounds[1],image.geometricBounds[0]]);
            }
            else
            {
                var imageGroupPlacedTextFrame=findPlacedTextFrame (imageGroup);
                var nextTextToImageGroupFrame=imageGroupPlacedTextFrame.nextTextFrame;
                var captionFrame=imageGroup.textFrames[0];
                captionFrame.move([nextTextToImageGroupFrame.geometricBounds[1],imageGroup.geometricBounds[0]]);
            }
        }
    }

    function checkPlacedOutsidePageMargin(imageGroup)
    {
        var isOutSidePageMargin=false;
        var imageGroupPage=imageGroup.parentPage;
        var pageMarginValue=myDoc.documentPreferences.pageHeight-imageGroupPage.marginPreferences.bottom;
        var bottomCoordinateOfImageGroupValue=imageGroup.geometricBounds[2];
        if(Math.round(pageMarginValue)<Math.round(bottomCoordinateOfImageGroupValue))
            isOutSidePageMargin=true;
        return isOutSidePageMargin;
    }

    function mapPGStyleNameToTemplateStylePrefix(styleName)
    {
        var prefixTemplateName="";
        switch(styleName)
        {
            case "Springervienna": prefixTemplateName="SV";
            case "GlobalJournalLarge":prefixTemplateName="GL";
        }
        return prefixTemplateName;
    }

    function getMasterPage(spreadName)
    {
        var returnMasterPage=null;
        var masterPagesLength = myDoc.masterSpreads.length
        for(var temp = masterPagesLength-1; temp>-1; temp--)
        {        
            var masterPage = myDoc.masterSpreads[temp];
            var masterPageName =masterPage.name; 
            var matchString=spreadName;
            if(masterPageName.indexOf(matchString)!=-1)
            {
                returnMasterPage=masterPage ;
                break;
            }
        }
        return returnMasterPage; 
    }

    function detachAllElementFromPage(myPage)
    {
        for (var j = myPage.appliedMaster.pageItems.length-1; j>0;  j--) 
        {
            try{
                if(myPage.appliedMaster.pageItems[j].label!="")
                {
                    if(myPage.appliedMaster.pageItems[j].label=="right_article_category"||myPage.appliedMaster.pageItems[j].label=="right_running_footer" ||myPage.appliedMaster.pageItems[j].label=="logo" || myPage.appliedMaster.pageItems[j].label == "left_running_title" || myPage.appliedMaster.pageItems[j].label == "right_running_title" ||myPage.appliedMaster.pageItems[j].label=="Society logo" )
                        continue;
                    else
                    {
                        myPage.appliedMaster.pageItems[j].override(myPage);
                        myPage.appliedMaster.pageItems[j].detach();
                    }
                }    
            } catch(e) {}
        }
    }

    function removeAQEmptyFrame(myDoc)
    {
        var aqChapter = getNodes(myDoc, ".//AQChapter");
        if(aqChapter.length > 0){
            detachFrameWithlabel ("AuthorQuery",  myDoc.pages.lastItem())
            addBlankPage = myDoc.pages.lastItem();
            var allTextFrames=addBlankPage.textFrames;
            for(var t=0;t<allTextFrames.length;t++)
            {
                if(allTextFrames[t].label.toLowerCase()=="authorquery")
                AQFrame= allTextFrames[t];
                if(AQFrame.associatedXMLElement==null && AQFrame.contents==""){
                    AQFrame.remove();
                    }
                }                
            }
        }


    function detachFrameWithlabel(frameLabel,myPage)
    {
        for (var j = myPage.appliedMaster.pageItems.length-1; j>0;  j--) 
        {
            try{
                if(myPage.appliedMaster.pageItems[j].label.toLowerCase()==frameLabel.toLowerCase())
                {
                    myPage.appliedMaster.pageItems[j].override(myPage);
                    myPage.appliedMaster.pageItems[j].detach();
                }    
            } catch(e) {}
        }
    }

    function findTextFrameWithLabel(framePage,frameLabel)
    {
        var labelTextFrame=null;
        var textFrames=framePage.textFrames;
        for(var t=textFrames.length-1;t>=0;t--)
        {
            var textFrame=textFrames[t]
            if(textFrame.label==frameLabel)
                labelTextFrame=textFrame;
        }
        return labelTextFrame;
    }

    function MoveElement(myxml, myStyleName)
    {
        var myProject=myxml.control;
        var myPageItems = myDoc.allPageItems;
        var myPage = myDoc.pages.everyItem();
        for(var i=0; i<myProject.length(); i++)
        {
            if(myProject[i].@StyleName==myStyleName)
            {
                var Xml_Log1=myProject[i].Element.*
                for (var e=0; e<Xml_Log1.length(); e++)
                {
                    var myElement = Xml_Log1[e].@Element;
                    var myAttributes = Xml_Log1[e].@type;
                    var moveAttributtes= Xml_Log1[e].@ismovable;
                    var isNoFitAttributes=Xml_Log1[e].@isnofit;
                    var fitAttributes= Xml_Log1[e].@isfit;
                    var deleteAttributes= Xml_Log1[e].@isdelete;
                    var isNoFit=false;
                    var isMovable=false;
                    var isFit=false;
                    var isDelete=false;
                    if(moveAttributtes!="")
                        isMovable=moveAttributtes;
                    if(isNoFitAttributes!="")
                        isNoFit=isNoFitAttributes;
                    if(fitAttributes!="")
                        isFit=fitAttributes;
                    if(deleteAttributes!="")
                        isDelete=deleteAttributes;
                    var xpath ="//"+myElement+"[@"+"type"+"="+"'"+ myAttributes+"'"+"]";
                    var root  = myDoc.xmlElements[0];
                    var node  = null;
                    try{
                        var proc  = app.xmlRuleProcessors.add([xpath]);
                        var match = proc.startProcessingRuleSet(root);
                        while( match!=undefined ) 
                        {
                            node = match.element;
                            var content=node.contents;
                            var myString=content.toString();
                            match = proc.findNextMatch();
                        }
                        for (var k = 0;  myDoc.pages.length>k; k++) 
                        {
                            myPage = myDoc.pages[k];
                            for (var r = 0; myPage.pageItems.length>r;  r++) 
                            {
                                var pageItem=myPage.pageItems[r];
                                try{
                                    if(pageItem.label ==myAttributes&&!isMovable && node!=null)
                                    {
                                        pageItem.placeXML(node);
                                        refresh();
                                        if(!isNoFit)
                                            pageItem.fit(FitOptions.FRAME_TO_CONTENT);
                                        if(isFit)
                                        {
                                          //Changes For three coloum Layout 
                                            var numberOfColumn=myPage.marginPreferences.columnCount;
                                            if(numberOfColumn == 3)
                                            {
                                                   var startCoordinate=pageItem.geometricBounds[1];
                                                   pageItem.move([-60,0]);
                                                   pageItem.fit(FitOptions.frameToContent);
                                                   pageItem.move([startCoordinate,200])
                                                   myDoc.align([pageItem],AlignOptions.BOTTOM_EDGES,AlignDistributeBounds.MARGIN_BOUNDS);
                                                   
                                            }
                                            else if (pageItem.label.indexOf("abstract")!=0)
                                            {
                                                    var placeTextFrame=findPlacedTextFrame(pageItem);
                                                    var objWidth=pageItem.geometricBounds[3]-pageItem.geometricBounds[1];
                                                    pageItem.move([pageItem.geometricBounds[1]-objWidth,pageItem.geometricBounds[1]]);
                                                    myDoc.align([pageItem],AlignOptions.TOP_EDGES,AlignDistributeBounds.MARGIN_BOUNDS);
                                                    pageItem.geometricBounds=[pageItem.geometricBounds[0],pageItem.geometricBounds[1],placeTextFrame.geometricBounds[2],pageItem.geometricBounds[3]];
                                                    pageItem.fit(FitOptions.FRAME_TO_CONTENT);
                                                    myDoc.align([pageItem],AlignOptions.LEFT_EDGES,AlignDistributeBounds.MARGIN_BOUNDS);
                                                    myDoc.align([pageItem],AlignOptions.BOTTOM_EDGES,AlignDistributeBounds.MARGIN_BOUNDS);
                                                    adjustAuthorCollectionFrame();
                                                    placeTextFrame.geometricBounds=[placeTextFrame.geometricBounds[0],placeTextFrame.geometricBounds[1],pageItem.geometricBounds[0],placeTextFrame.geometricBounds[3]];
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if(isDelete&&pageItem.label ==myAttributes&&!isMovable)
                                        {
                                            if(pageItem.contents=="")
                                            {
                                                pageItem.remove();
                                            }
                                         }
                                    }
                                }catch(e){
                                    alert(e)
                                }
                            }
                        }
                    }catch( ex ){
                        alert(ex);
                    }finally{
                        proc.endProcessingRuleSet();
                        proc.remove();
                    }
                }
            }
        }
    }

    
    function MoveElementForRunningHead(myxml, myStyleName)
    {
        logMsg = logMsg + (++msgCounter) + ". Move element for running head. " + "\n"
        var myProject=myxml.control;
        var myPageItems = myDoc.allPageItems;
        var myPage = myDoc.pages.everyItem();
        for(var i=0; i<myProject.length(); i++)
        {
            if(myProject[i].@StyleName==myStyleName)
            {
                var Xml_Log1=myProject[i].Element.*
                for (var e=0; e<Xml_Log1.length(); e++)
                {
                    myAttributeCntr = 0;
                    var myElement = Xml_Log1[e].@Element;
                    var myAttributes = Xml_Log1[e].@type;
                    var moveAttributtes= Xml_Log1[e].@ismovable;
					var deleteAttributes= Xml_Log1[e].@isdelete;
                    var isMovable=false;
					var isDelete = false;
                    if(moveAttributtes!="")
						isMovable=moveAttributtes;
					if(deleteAttributes!="")
                        isDelete = deleteAttributes;
                    var xpath ="//"+myElement+"[@"+"type"+"="+"'"+ myAttributes+"'"+"]";
                    var root  = myDoc.xmlElements[0];
                    var node  = null;
                    try{
                        var proc  = app.xmlRuleProcessors.add([xpath]);
                        var match = proc.startProcessingRuleSet(root);
                        while( match!=undefined ) 
                        {
                            node = match.element;
                            var content=node.contents;
                            var myString=content.toString();
                            match = proc.findNextMatch();
                        }
                        for (var k = 0;  k<myDoc.pages.length; k++) 
                        {
                            myPage = myDoc.pages[k];
                            for (var r = 0; myPage.appliedMaster.pageItems.length>r;  r++) 
                            {
                                var pageItem=myPage.appliedMaster.pageItems[r];
                                try{
                                    if(isMovable)
                                    {
                                        if(pageItem.label ==myAttributes)
                                        {
                                            if(myAttributes=="right_running_title"){
                                                if(myAttributeCntr == 0){
                                                    node.xmlAttributes.add ("aid:pstyle", "Running_Head_Recto");                                               
                                                    pageItem.placeXML(node);
                                                    myAttributeCntr++;
                                                    }
                                                }
                                            else if(myAttributes=="left_running_title"){
                                                if(myAttributeCntr == 0){
                                                    node.xmlAttributes.add ("aid:pstyle", "Running_Head_Verso");                                               
                                                    pageItem.placeXML(node);
                                                    myAttributeCntr++;
                                                    }
                                                }
                                            else{
                                            if(node!=null)
                                                changeRunningHeadTextForJournal(pageItem, node.contents);
                                            else
                                                changeRunningHeadTextForJournal(pageItem, "");
                                                }
                                        }
                                    }
									if(isDelete && pageItem.label ==myAttributes && node==null)
                                    {
                                        if(pageItem.contents=="")
                                        {
                                            pageItem.remove();
                                        }
                                    }
                                }catch(e) {alert(e)}
                            }
                        }  
                    }catch( ex ){
                        alert(ex);
                    }finally{
                        proc.endProcessingRuleSet();
                        proc.remove();
                    }
                }
            }
        }
    }


    function detchAllHeadersAndFootersFromAllPages()
    {
        var pages=myDoc.pages;
        for(var i=0;i<pages.length;i++)
        {
            detachAllElementFromPage(myDoc.pages[i]);
        }
    }

    function getNodesForJournals()
    {
        var objectElement=new Array();
        var xPath = "//cs_repos"
        var root  = myDoc.xmlElements[0];
        var node  = null;
        try{
            var proc  = app.xmlRuleProcessors.add([xPath, ]);
            var match = proc.startProcessingRuleSet(root);
            while( match!=undefined ) 
            { 
                node = match.element;
                var xmlAttribute=node.xmlAttributes.item("typename");
                if(xmlAttribute.isValid)
                {
                    var xmlAttributeValue=xmlAttribute.value;
                    if(xmlAttributeValue.toString().toLowerCase()=="table")
                        objectElement.push(node)
                    else if(xmlAttributeValue.toString().toLowerCase()=="figure")
                        objectElement.push(node)
                    else if(xmlAttributeValue.toString().toLowerCase()=="query")
                        objectElement.push(node)    
                }
                match = proc.findNextMatch(); 
            }
        }catch( ex ){
            alert(ex);
        }
        finally{
            proc.endProcessingRuleSet();
            proc.remove();
        }
        return objectElement;
    }
    
    function changeSwatchForLayout()
    {
        logMsg = logMsg + (++msgCounter) + ". Change swatch for layout. " + "\n"
        var dummySwatchName="Dummy";
        var particularSwatch=swatchByName(prefixNumber);
        if(particularSwatch!=null)
        {
            var allSwatches=myDoc.swatches; 
            for(var s=0;s<allSwatches.length;s++)
            {
                if(allSwatches[s].name.toLowerCase()==dummySwatchName.toLowerCase())
                {
                    allSwatches[s].remove(particularSwatch);
                    break;
                }
            }
        }    
    }

    function swatchByName(swatchName)
    {
        var particularSwatch=null;
        var allSwatches=myDoc.swatches; 
        for(var s=0;s<allSwatches.length;s++)
        {
            if(allSwatches[s].name.toLowerCase()==swatchName.toString().toLowerCase())
                particularSwatch=allSwatches[s];
        }
        return particularSwatch;
    }

    function refreshFigSideCaptionStyle(myDoc)
    {
        logMsg = logMsg + (++msgCounter) + ". Refresh Figure Side Caption Styles. " + "\n"
        var allGroups=myDoc.groups;
        for(var g=0;g<allGroups.length;g++)
        {
            var group=allGroups[g];
            applySideCaptionStyle (group)
            var groups=group.groups;
            for(var gr=0;gr<groups.length;gr++)
            {
                applySideCaptionStyle (groups[gr]);
            }
        }
    }

    function applySideCaptionStyle(group)
    {
        var groupTextFrames=group.textFrames;
        for(var t=0;t<groupTextFrames.length;t++)
        {
            var groupTextFrame=groupTextFrames[t]
            if(groupTextFrame.label.toLowerCase()=="sidecaptionframe")
                applyStyleToSideCaptionTextFrame(groupTextFrame);
            if(groupTextFrame.label.toLowerCase()=="tablesidecaption")
                applyStyleToSideTableTextFrame(groupTextFrame);
            if(groupTextFrame.label.toLowerCase()=="tablesidefooter")
                applyStyleToFootnoteTableTextFrame(groupTextFrame);
        }
    }

    function applyStyleToSideCaptionTextFrame(captionTextFrame)
    {
        try{
            captionTextFrame.insertionPoints.everyItem().paragraphs.everyItem().appliedParagraphStyle=myDoc.paragraphStyles.item("Figure_Caption_Side");
            captionTextFrame.fit(FitOptions.FRAME_TO_CONTENT);
        }catch(e){}
    }

    function applyStyleToSideTableTextFrame(tableTextFrame)
    {
        try{
            if(myTemplateJournalID=="00501")
            {
                tableTextFrame.insertionPoints.everyItem().paragraphs.everyItem().appliedParagraphStyle=myDoc.paragraphStyles.item("Table_Caption_Side");
                tableTextFrame.fit(FitOptions.FRAME_TO_CONTENT);
                tableTextFrame.insertionPoints.everyItem().paragraphs.firstItem().appliedParagraphStyle=myDoc.paragraphStyles.item("Table_Number_Side");
                tableTextFrame.fit(FitOptions.FRAME_TO_CONTENT);
            }
            else
            {
                tableTextFrame.insertionPoints.everyItem().paragraphs.everyItem().appliedParagraphStyle=myDoc.paragraphStyles.item("Table_Caption_Side");
                tableTextFrame.fit(FitOptions.FRAME_TO_CONTENT);
            }
        }catch(e){alert()}
    }

    function applyStyleToFootnoteTableTextFrame(tableTextFrame)
    {
        try{
            var oldVariable = tableTextFrame.geometricBounds[2];
            tableTextFrame.insertionPoints.everyItem().paragraphs.everyItem().appliedParagraphStyle=myDoc.paragraphStyles.item("Table_Footnote");
            tableTextFrame.fit(FitOptions.FRAME_TO_CONTENT);
            var newVariable = tableTextFrame.geometricBounds[2];
            var diff=newVariable -oldVariable ;
            tableTextFrame.move([tableTextFrame.geometricBounds[1],tableTextFrame.geometricBounds[0]-diff]);
        }catch(e){}
    }

    function getColumnMapValue(myTemplateColumns )
    {
        var layoutName="";
        var col=myTemplateColumns.toString();
        switch( col)
        {
            case "1 Column":layoutName="1_Col_Journals.indt" ;
                        break;
            case "2 Column":layoutName="2_Col_Journals.indt";
                        break;
            case "3 Column":layoutName="3_Col_Journals.indt" ;
                        break;
            case "4 Column":layoutName="3_Col_Journals.indt" ;
                        break;            
        } 
        return layoutName;
    }

    function getTableColumnCountFromXmlElement(tableElement)
    {
        var count=0;
        var tableAttributes=tableElement.xmlAttributes;
        for(var i=0;i<tableAttributes.length;i++)
        {
            if(tableAttributes[i].name.toLowerCase().indexOf("cs_col")!=-1)
                count++;
        }
        return count;
    }

    function findAndPlaceLogo()
    {
        var logoFolderPath=File(appPath +"/Logo");
        var  myLogoPath = logoFolderPath.getFiles("*.*");
        for(var p=0;p<myLogoPath.length;p++)
        {
            var logo=myLogoPath[p].toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
            var logoName=new File(logo).name; 
            
            //Vivek: for placing Logos.
           logoNumber = logoName.toString().split("_").reverse().pop();
           if(logoNumber == prefixNumber)
           {
                 PlaceLogo(myDoc,logo);
            }
        }
    }

    function PlaceLogo(myDoc,logoPath)
    {
        for(var i=0;i<myDoc.pages.length;i++)
        {
            var page=myDoc.pages[i];
            var allTextFrames=page.appliedMaster.pageItems;
            for(var j=0;j<allTextFrames.length;j++)
            {
                if(allTextFrames[j].label=="logo")
                {
                    try{
                        var rec=allTextFrames[j].place(logoPath);
                        myDoc.align([rec[0]],AlignOptions.RIGHT_EDGES,AlignDistributeBounds.MARGIN_BOUNDS)
                        var recHeight=allTextFrames[j].geometricBounds[2]-allTextFrames[j].geometricBounds[0];
                        var recImageHeight=rec[0].geometricBounds[2]-rec[0].geometricBounds[0];
                        if(parseInt (recImageHeight)>parseInt(recHeight))
                            rec[0].fit(FitOptions.FRAME_TO_CONTENT);
                        myDoc.align([rec[0]],AlignOptions.RIGHT_EDGES,AlignDistributeBounds.MARGIN_BOUNDS)
                        rec[0].itemLink.unlink();
                    }catch(e){}
                }
            }
        }
    }


 function findAndPlaceSocietyLogo()
    {
        var logoFolderPath=File(appPath +"/Logo/Society logo");
        var  myLogoPath = logoFolderPath.getFiles("*.*");
        for(var p=0;p<myLogoPath.length;p++)
        {
            var logo=myLogoPath[p].toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
            var logoName=new File(logo).name; 
            
            //Vivek: for placing Logos.
           logoNumber = logoName.toString().split("_").reverse().pop();
           if(logoNumber == prefixNumber)
           {
                 PlaceSocietyLogo(myDoc,logo);
            }
        }
    }

 function PlaceSocietyLogo(myDoc,logoPath)
    {
        for(var i=0;i<myDoc.pages.length;i++)
        {
            var page=myDoc.pages[i];
            var allTextFrames=page.appliedMaster.pageItems;
            for(var j=0;j<allTextFrames.length;j++)
            {
                $.writeln(allTextFrames[j].label);
                if(allTextFrames[j].label=="Society logo")
                {
                    try{
                                var rec=allTextFrames[j].place(logoPath);
                                var recHeight=allTextFrames[j].geometricBounds[2]-allTextFrames[j].geometricBounds[0];
                                var recWidth=allTextFrames[j].geometricBounds[3]-allTextFrames[j].geometricBounds[1];
                                var recImageHeight=rec[0].geometricBounds[2]-rec[0].geometricBounds[0];
                                var recImageWidth=rec[0].geometricBounds[3]-rec[0].geometricBounds[1];

                              var xCo = allTextFrames[j].geometricBounds[3] - recImageWidth;
                              var yCo = allTextFrames[j].geometricBounds[2] - recImageHeight;
                               
                                   try{
                                   rec[0].move([xCo, yCo]);
                                   rec[0].fit(FitOptions.FRAME_TO_CONTENT);
                                   rec[0].itemLink.unlink();
                                   }
                                   catch(e)
                                   {
                                       alert(e);
                                   }
                                  
                                if(parseInt (recImageHeight)>parseInt(recHeight))
                                    rec[0].fit(FitOptions.FRAME_TO_CONTENT);
                                
                        }catch(e){}
                }
            }
        }
    }


    function resizeTableBeforeSave(myDoc)
    {
        for (s=0; s<myDoc.stories.length; s++)
        {
            for (t=0; t<myDoc.stories[s].tables.length; t++)
            {
                try{
                    check(myDoc.stories[s].tables[t],myDoc);
                }catch(ex){ }
            }     
        }

        function check(aTable,myDoc)
        {
            var tableElement=aTable.associatedXMLElement.parent.parent;
            var tableStyle=tableElement.xmlAttributes.item("Style").value ; 
            var colwidth=null;
            var cols=aTable.columns;
            var  singelTextFrameWidth=myDoc.documentPreferences.pageWidth-myDoc.pages[2].marginPreferences.left-myDoc.pages[2].marginPreferences.left;
            var textFrameHeight=myDoc.documentPreferences.pageHeight-myDoc.pages[2].marginPreferences.top-myDoc.pages[2].marginPreferences.bottom
            colwidth=Math.round(textFrameHeight)/Math.round(aTable.columns.length);
            for(var i=0;i<cols.length;i++)
            {
                aTable.columns[i].width = colwidth;
            }
        }
    }
    
    function floatingPlacement()
    {
        if(TemplateName=="3_Col_Journals.indt")
        {
            try{
                fazFloatPlacement();
                _gOplObj.startObjectPlacement();
            }catch(e){}
            
        }
        else
        {
            imageAndTablePlacement();
        }
    }


    function fazFloatPlacement()
    {
        try{
            var myScriptPath = File (sAppPath+"/"+"Scripts/Breeze/FigTabPlacement.jsxbin");
            myScriptPath = myScriptPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
            Fpath=File(myScriptPath);
            myRefresh = Fpath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
            app.doScript(myRefresh);
       }catch(e){}
    }

    
    function imageAndTablePlacement()
    {
        logMsg = logMsg + (++msgCounter) + ". Image and Table Placement. " + "\n"
        var objectElement=new Array();
        var xPath = "//cs_repos"
        var root  = myDoc.xmlElements[0];
        var node  = null;
        var flag=true;
        try{
            var proc  = app.xmlRuleProcessors.add([xPath, ]);
            var match = proc.startProcessingRuleSet(root);
            while( match!=undefined ) 
            { 
                node = match.element;
                var xmlAttribute=node.xmlAttributes.item("typename");
                if(xmlAttribute.isValid)
                {
                    var xmlAttributeValue=xmlAttribute.value;
                    if(xmlAttributeValue.toString().toLowerCase()=="table")
                        createAndPlaceImageAndTableInDocument(node);
                    else if(xmlAttributeValue.toString().toLowerCase()=="figure")
                        createAndPlaceImageAndTableInDocument(node)
                }
                match = proc.findNextMatch(); 
            }
        }catch( ex ){
            logMsg = logMsg + (++msgCounter) + ". ERROR in Image and Table Placement. " + "\n"+ex+"\n"
            alert(ex);
        }
        finally{
            proc.endProcessingRuleSet();
            proc.remove();
        }
        return objectElement;
    }

    function createAndPlaceImageAndTableInDocument(objectInternalRefNode)
    {
        var styleName="";
        var objectInternalRefAttribute=objectInternalRefNode.xmlAttributes.item("repoID");
        if(objectInternalRefAttribute.isValid)
        {
            var floatingRepoId=objectInternalRefAttribute.value;
            var floatingRepositionXmlElement=getNodesFromAttributeTypeAndValue("reposition","cs:repoID",floatingRepoId.toString());
            if(floatingRepositionXmlElement!=null)
            {
                var floatingXmlElement=floatingRepositionXmlElement.xmlElements[0];
                var floatingType=floatingElementType(objectInternalRefNode);
                var styleIdAttribute=floatingXmlElement.xmlAttributes.item("Style");
                if(styleIdAttribute.isValid&&floatingType.toLowerCase()=="table")
                    styleName=styleIdAttribute.value;
                var floatingIdAttribute=floatingXmlElement.xmlAttributes.item("ID");
                $.writeln (floatingIdAttribute.value)
                var floatingTypeAttribute=floatingXmlElement.xmlAttributes.item("Float");
                if(floatingIdAttribute.isValid && floatingTypeAttribute.isValid)
                {
                    var typeFloat=floatingTypeAttribute.value;
                    if(typeFloat.toLowerCase()!="no")
                    {
                        var floatingId=floatingIdAttribute.value;
                        var floatingInternalRef= getNodesFromAttributeTypeAndValue("InternalRef","RefID",floatingId.toString());
                        if(floatingInternalRef!=null)
                        {
                            var floatingPage=floatingInternalRef.insertionPoints[0].parentTextFrames[0].parentPage;  
                            objectPlacement (objectInternalRefNode, floatingPage,floatingType,false)
                        }
                        else
                        {
                            var endPage=myDoc.pages.add(LocationOptions.AT_END);
                            objectPlacement (objectInternalRefNode, endPage,floatingType,true);
                        }
                    }
                }   
            }
        }
    }
    function objectPlacement(objectInternalRefNode,floatingPage,floatingType,isAtEnd)
    {
        var imageCaptionGroup=null;
        if(floatingType.toLowerCase()=="figure")
            imageCaptionGroup=makeImage(objectInternalRefNode,floatingPage)
        else if(floatingType.toLowerCase()=="table")    
            imageCaptionGroup=makeTable(objectInternalRefNode,floatingPage)
        if(imageCaptionGroup!=null/*&&!isAtEnd*/)
            var objectCaptionGroup=placeObject(imageCaptionGroup,objectInternalRefNode);
    }


    function makeImage(objectInternalRefNode,imagePage)
    {
        var imageCaptionGroup=null;
        var objectInternalRefAttribute=objectInternalRefNode.xmlAttributes.item("repoID");
        if(objectInternalRefAttribute.isValid)
        {
            var imageRepoId=objectInternalRefAttribute.value;
            var imageRepositionXmlElement=getNodesFromAttributeTypeAndValue("reposition","cs:repoID",imageRepoId.toString());
            if(imageRepositionXmlElement!=null)
            {
                var figureXmlElement=imageRepositionXmlElement.xmlElements[0];
                var figurePathXmlAttribute=figureXmlElement.xmlAttributes.item("cs_Fignode");
                if(figurePathXmlAttribute.isValid)
                {
                    var imagePath=figurePathXmlAttribute.value;
                    var isImageFileExist=checkFileIsExist(imagePath);
                    if(!isImageFileExist)
                    {
                        imagePath="D:/FigNotFound.jpg";
                    }
                    var xImageCoordinate=0;
                    var articleNode=getNodesFromAttributeTypeAndValue("Article","","");
                    var articleStartPage=articleNode.insertionPoints[0].parentTextFrames[0].parentPage;
                    if(articleStartPage.isValid)
                    {
                        if(parseInt(articleStartPage.name)>=parseInt(imagePage.name))
                            imagePage=nextPage (imagePage)
                    }
                    if(lastImagePlacedPage!=null)
                    {
                        if(parseInt(lastImagePlacedPage.name)>parseInt(imagePage.name))
                            imagePage=lastImagePlacedPage;
                    }
                    var topObject=topPlacementObjectWithinSamePage(imagePage);
                    var yImageCoordinate=imagePage.marginPreferences.top;
                    if(parseInt(imagePage.name)%2==0)
                        xImageCoordinate=imagePage.marginPreferences.right
                    else if(parseInt(imagePage.name)%2!=0)
                        xImageCoordinate=imagePage.marginPreferences.left
                    var topObject=topPlacementObjectWithinSamePage(imagePage);
                    var imageRectangle=imagePage.place(imagePath,[xImageCoordinate,yImageCoordinate]);
                    imageRectangle[0].parent.appliedObjectStyle = myDoc.objectStyles.item("FigureStyle");
                    if(topObject!=null)
                    {
                        var topImageType=typeOfImage(topObject);
                        var imageType=typeOfImage(imageRectangle[0]);
                        var imageHeight = imageRectangle[0].geometricBounds[2] - imageRectangle[0].geometricBounds[0];
                        if(imageType=="S"&&topImageType=="S")
                        {
                            var topFullWidthObject=getTopFullWidthObject (imagePage);
                                yImageCoordinate=topObject.geometricBounds[2]+8;
                        }
                        else
                            yImageCoordinate=topObject.geometricBounds[2]+8;
                        if(imageType=="S"&&topSingleWidth!=null)
                        {
                            if(topSingleWidth.parentPage.id==imagePage.id)
                                xImageCoordinate=topSingleWidth.geometricBounds[1];
                        }
                        imageRectangle[0].parent.move([xImageCoordinate, yImageCoordinate]);
                        myDoc.save(SavePath); 
                    }
                    var pageHeight= myDoc.documentPreferences.pageHeight-imagePage.marginPreferences.bottom;
                    var remainSps = Math.round(pageHeight) - Math.round(yImageCoordinate);
                    if(Math.round(yImageCoordinate)>=Math.round(pageHeight))
                    {
                        imagePage=nextPage(imagePage)
                        yImageCoordinate=imagePage.marginPreferences.top;
                    }
                    else if(imageHeight > remainSps)
                    {
                        imagePage=nextPage(imagePage)
                        yImageCoordinate=imagePage.marginPreferences.top;
                    }
                
                    imageRectangle[0].parent.move(imagePage);
                    imageRectangle[0].parent.move([xImageCoordinate, yImageCoordinate]);       
                    imageStyle = getImageStyle (imageRectangle[0].parent)
                    var captionNode=getNodesFromAttributeTypeAndValue ("cs_CaptionWrapper/Caption", "", "", figureXmlElement)
                    if(captionNode!=null)
                     {
                        var parentTag = captionNode.parent.parent.xmlAttributes.item("ID").value; 
                        $.writeln (parentTag)
                        var splitedIdTag = parentTag.toString().substring (0, 3);
                        captionTextFrame = createCaptionForImage(captionNode, imageRectangle[0],imageRectangle[0].parentPage);  
                        var imageCaptionGroup=imagePage.groups.add([imageRectangle[0].parent,captionTextFrame])
                        imageCaptionGroup.label=parentTag;
                    }
                                
                }
            }
        }
        return imageCaptionGroup;
    }
    
    
    function getImageStyle(imageFrame){
    imageWidth = parseInt(imageFrame.geometricBounds[3]) - parseInt(imageFrame.geometricBounds[1]);
    if(myTemplateName == "Large_Template"){
        if(imageWidth > 0 && imageWidth <= 39){
             return "STYLE_1";
            }
        else if(imageWidth > 39 && imageWidth <= 84){
             return "STYLE_2";
            }
        else if(imageWidth > 84 && imageWidth <= 129){
             return "STYLE_3";
            }
        else if(imageWidth > 129 && imageWidth <= 174){
             return "STYLE_4";
            }
        else{
             return "STYLE_5";
            }        
        }
    else if(myTemplateName == "Medium_Template"){
        if(imageWidth > 0 && imageWidth <= 34){
             return "STYLE_1";
            }
        else if(imageWidth > 34 && imageWidth <= 76){
             return "STYLE_2";
            }
        else if(imageWidth > 76 && imageWidth <= 118){
             return "STYLE_3";
            }
        else if(imageWidth > 118 && imageWidth <= 160){
             return "STYLE_4";
            }
        else{
             return "STYLE_5";
            }        
        }
    else if(myTemplateName == "Small_Condensed_Template"){
        if(imageWidth > 0 && imageWidth <= 80){
             return "STYLE_3";
            }
        else if(imageWidth > 80 && imageWidth <= 122){
             return "STYLE_4";
            }
        else{
             return "STYLE_5";
            }        
        }
    else if(myTemplateName == "Small_Extended_Template"){
        if(imageWidth > 0 && imageWidth <= 78){
             return "STYLE_3";
            }
        else if(imageWidth > 78 && imageWidth <= 119){
             return "STYLE_4";
            }
        else{
             return "STYLE_5";
            }        
        }
    }

    function makeTable(objectInternalRefNode,tablePage)
    {
        var tableTextFrame=null;
        var objectInternalRefAttribute=objectInternalRefNode.xmlAttributes.item("repoID");
        if(objectInternalRefAttribute.isValid)
        {
            var tableRepoId=objectInternalRefAttribute.value;
            var tableRepositionXmlElement=getNodesFromAttributeTypeAndValue("reposition","cs:repoID",tableRepoId.toString());
            if(tableRepositionXmlElement!=null)
            {
                if(lastImagePlacedPage!=null)
                {
                    if(parseInt(lastImagePlacedPage.name)>parseInt(tablePage.name))
                        tablePage=lastImagePlacedPage;
                }
            }
            var pageHeight=myDoc.documentPreferences.pageHeight;
            var pageWidth=myDoc.documentPreferences.pageWidth;
            var left=tablePage.marginPreferences.left;
            var right=tablePage.marginPreferences.right;
            var top=tablePage.marginPreferences.top;
            var bottom=tablePage.marginPreferences.bottom;
            var tableXmlElement=tableRepositionXmlElement.xmlElements[0];
            var tableStyle="";
            var tableStyleAttribute=tableXmlElement.xmlAttributes.item("Style");
            if(tableStyleAttribute.isValid)
                tableStyle=tableStyleAttribute.value;
            var tableElement=getNodesFromAttributeTypeAndValue ("Story/Table", "", "", tableXmlElement);    
            if(tableElement!=null)
            {
                expandTable(tableElement,tablePage);
                var tables=tableElement.tables;
                var xImageCoordinate=0;
                var yImageCoordinate=top;
                var topObject=topPlacementObjectWithinSamePage(tablePage);
                if(topObject!=null)
                {
                    var isNextPage=false;
                    if(tables.length>0)
                    {
                        var tableHeight=tables[0].height;
                        var bottomPlacedObject=bottomPlacementObjectWithinSamePage(tablePage);
                        var topPlacedObject=topPlacementObjectWithinSamePage(tablePage);
                        var remainSpace=myDoc.documentPreferences.pageHeight-tablePage.marginPreferences.top-tablePage.marginPreferences.bottom;
                        if(topPlacedObject==null&&bottomPlacedObject!=null)
                            remainSpace=myDoc.documentPreferences.pageHeight-tablePage.marginPreferences.top-(bottomPlacedObject.geometricBounds[2]-bottomPlacedObject.geometricBounds[0])-8;
                        else if(topPlacedObject!=null&&bottomPlacedObject==null)
                            remainSpace=myDoc.documentPreferences.pageHeight-(topPlacedObject.geometricBounds[2]+8)-tablePage.marginPreferences.bottom;
                        else if(topPlacedObject!=null&&bottomPlacedObject!=null)
                            remainSpace=(bottomPlacedObject.geometricBounds[0]-8)-(topPlacedObject.geometricBounds[2]+8)
                        if(Math.round(remainSpace)<Math.round(tableHeight))
                            isNextPage=true;
                    }
                    if(tableStyle.toLowerCase()=="style4"||isNextPage)
                    {
                        tablePage=nextPage(tablePage);
                        pageHeight=myDoc.documentPreferences.pageHeight;
                        pageWidth=myDoc.documentPreferences.pageWidth;
                        left=tablePage.marginPreferences.left;
                        right=tablePage.marginPreferences.right;
                        top=tablePage.marginPreferences.top;
                        bottom=tablePage.marginPreferences.bottom;
                        tableFrameHeight=pageHeight-top;
                    }
                    else
                    {
                        yImageCoordinate=topObject.geometricBounds[2]+8;
                        tableFrameHeight=pageHeight-bottom;
                    }    
                }
                var tableFrameWidth=pageWidth-left;
                var tableFrameHeight=pageHeight-bottom;
                if(parseInt(myDoc.pages[0].name)>=parseInt(tablePage.name))
                    tablePage=nextPage(tablePage);
                if(tablePage.side==PageSideOptions.LEFT_HAND)
                {
                    xImageCoordinate=tablePage.marginPreferences.right;
                    tableFrameWidth=pageWidth-left
                }
                else if(tablePage.side==PageSideOptions.RIGHT_HAND)
                {
                    xImageCoordinate=tablePage.marginPreferences.left
                    tableFrameWidth=pageWidth-right;
                }
                if(tableStyle.toLowerCase()=="style1"&&topSingleWidth!=null)
                {
                    if(topSingleWidth.parentPage.id==tablePage.id)
                        xImageCoordinate=topSingleWidth.geometricBounds[1];
                }
                tableTextFrame=tablePage.textFrames.add();
                tableTextFrame.label="Table";
                tableTextFrame.geometricBounds=[yImageCoordinate,xImageCoordinate,tableFrameHeight,tableFrameWidth];
                if(tableStyle.toLowerCase()=="style4")
                {
                    var frameWidth=tableTextFrame.geometricBounds[3]-tableTextFrame.geometricBounds[1];
                    var frameHeight=tableTextFrame.geometricBounds[2]-tableTextFrame.geometricBounds[0];
                    tableTextFrame.geometricBounds=[tableTextFrame.geometricBounds[0],tableTextFrame.geometricBounds[1],tableTextFrame.geometricBounds[0]+frameWidth,tableTextFrame.geometricBounds[1]+frameHeight];
                }    
                tableTextFrame.placeXML(tableXmlElement);
                tableTextFrame.fit(FitOptions.frameToContent);
                 var captionTag = tableTextFrame.associatedXMLElement;
                var parentTag = captionTag.parent;
                   if(!parentTag.xmlAttributes.item("ID").isValid){
                        tableTextFrame.appliedObjectStyle = myDoc.objectStyles.item("FigureStyle");//Apply TableStyle Separate.
                    }
                tableTextFrame.fit(FitOptions.frameToContent);
                /*******   Condition For Continuation table on nextPage            *****/                
                if(tableTextFrame.overflows)
                {
                    var allBottomPlacedObjects=getAllBottomPlacedObjectInPage(tablePage)
                    if(allBottomPlacedObjects.length>0)
                    {
                        tablePage=nextPage(tablePage)
                        tableTextFrame.remove();
                        tableTextFrame=makeTable(objectInternalRefNode, tablePage);
                        tableTextFrame.geometricBounds=[tableTextFrame.geometricBounds[0],tableTextFrame.geometricBounds[1],tableFrameHeight,tableTextFrame.geometricBounds[3]];
                        return tableTextFrame;
                    }
                }
                if(tableStyle.toString().toLowerCase()=="style2")
                {
                    myDoc.align([tableTextFrame],AlignOptions.RIGHT_EDGES,AlignDistributeBounds.MARGIN_BOUNDS)
                    tableTextFrame=placeTwoThirdCaptionTableForJournals (tableTextFrame, true);
                }
                if(tableStyle.toLowerCase()=="style4")
                {
                    myDoc.zeroPoint=[0,0];
                    tablePage = tableTextFrame.parentPage;
                    xCor = tableTextFrame.geometricBounds[1];
                    tableTextFrame.absoluteRotationAngle =90;
                    myDoc.zeroPoint=[0,0];
                    tableTextFrame.move(tablePage);
                    tableTextFrame.move([0, 0]);
                    tableTextFrame.move([xCor,pageHeight-bottom])
                }
            }
        }
        return tableTextFrame;
    }


    function checkFileIsExist(filePath)
    {
        var isExist=false;
        var file=new File(filePath);
        if(file.exists)
            isExist=true;
        return isExist;
    }

    function getNodesFromAttributeType(attribiuteName,startElement)
    {
        var xmlElements=new Array();
        var xPath = "//cs_repos"
        var root  = startElement;
        var node  = null;
        try{
            var proc  = app.xmlRuleProcessors.add([xPath, ]);
            var match = proc.startProcessingRuleSet(root);
            while( match!=undefined ) 
            { 
                node = match.element;
                var xmlAttribute=node.xmlAttributes.item("typename");
                if(xmlAttribute.isValid)
                    xmlElements.push(node);
                match = proc.findNextMatch(); 
            }
        }
        catch( ex ){
            alert(ex);
        }
        finally{
            proc.endProcessingRuleSet();
            proc.remove();
        }
        return xmlElements;
    }

    function getNodesFromAttributeTypeAndValue(elementName,attribiuteName,attribiuteValue,startElement)
    {
        var xmlElement=null;
        var xPath = "//"+elementName.toString();
        var root  = startElement;
        var node  = null;
        try{
            var proc  = app.xmlRuleProcessors.add([xPath, ]);
            if(root==undefined)
                root=myDoc.xmlElements[0];
            var match = proc.startProcessingRuleSet(root);
            while( match!=undefined ) 
            { 
                node = match.element;
                if(attribiuteName.toString()=="")
                {
                    xmlElement=node;
                    break;
                }
                var xmlAttribute=node.xmlAttributes.item(attribiuteName.toString());
                if(xmlAttribute.isValid)
                {
                    if(xmlAttribute.value.toString().toLowerCase()==attribiuteValue.toString().toLowerCase())
                    {
                        xmlElement=node;
                        break;
                    }
                }
                match = proc.findNextMatch(); 
            }
        }catch( ex ){
            alert(ex);
        }
        finally{
            proc.endProcessingRuleSet();
            proc.remove();
        }
        return xmlElement;
    }

    function createCaptionForImage(captionNode,imageRectangle,imagePage)
    {
        var captionTextFrame=null;
        var imageType=typeOfImage(imageRectangle)        
        var captionTextFrame=imagePage.textFrames.add();
        var left=imagePage.marginPreferences.left;
        var right=imagePage.marginPreferences.right;
        var top=imagePage.marginPreferences.top;
        var bottom=imagePage.marginPreferences.bottom;
        var pageHeight=myDoc.documentPreferences.pageHeight;
        var pageWidth=myDoc.documentPreferences.pageWidth;
        var marginWidth = pageWidth - (left + right);
        var marginHeight=pageHeight-top-bottom;
        captionTextFrame.label="figureCaptionFrame";;
        if(imageStyle=="STYLE_1")
        {
            
            var singleTextFrameWidth=getSingleTextFrameWidthFromPage(imagePage);
            if(myTemplateName == "Large_Template")
            {
                imageRectangle.parent.move([imageRectangle.geometricBounds[1]+(84-(imageRectangle.geometricBounds[3]-imageRectangle.geometricBounds[1])), imageRectangle.geometricBounds[0]])
                if(parseInt(imagePage.name)%2==0)
                    captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[0],right,imageRectangle.geometricBounds[2],right+39]
                else if(parseInt(imagePage.name)%2!=0)
                    captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[0],left,imageRectangle.geometricBounds[2],left+39]
            }
            else if(myTemplateName == "Medium_Template")
            {
                if(parseInt(imagePage.name)%2==0)
                    captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[0],right,imageRectangle.geometricBounds[2],right+34]
                else if(parseInt(imagePage.name)%2!=0)
                    captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[0],left,imageRectangle.geometricBounds[2],left+34]
            }
            captionTextFrame.placeXML(captionNode);
            captionTextFrame.appliedObjectStyle = myDoc.objectStyles.item("FigSideCaption");
            
            if(captionTextFrame.overflows){
                captionTextFrame.remove();
                captionTextFrame=changeCaptionFrameToSingleWidth(imageRectangle,captionNode);
                }
            else{
                   captionTextFrame.insertionPoints[0].paragraphs[0].appliedParagraphStyle=myDoc.paragraphStyles.item("Figure_Caption_Side");
                }
        }        
        else if(imageStyle=="STYLE_2")
        {
            var singleTextFrameWidth=getSingleTextFrameWidthFromPage(imagePage);
            if(parseInt(imagePage.name)%2==0)
                captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[2],right,imageRectangle.geometricBounds[2]+marginHeight,right+singleTextFrameWidth]
            if(parseInt(imagePage.name)%2!=0)
                captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[2],left,imageRectangle.geometricBounds[2]+marginHeight,left+singleTextFrameWidth]
            captionTextFrame.placeXML(captionNode);
            captionTextFrame.appliedObjectStyle = myDoc.objectStyles.item("FigureCaptionStyle_1Col");
            refresh();
            captionTextFrame.fit(FitOptions.frameToContent);
            myDoc.align([captionTextFrame,imageRectangle],AlignOptions.HORIZONTAL_CENTERS,AlignDistributeBounds.ITEM_BOUNDS);            
            imageRectangle.fit(FitOptions.frameToContent);
            var captionTag = captionTextFrame.associatedXMLElement;
            var parentTag = captionTag.parent.parent;
            if(parentTag.markupTag.name == "Figure"){
                var idTag = parentTag.xmlAttributes.item("ID").value;
                var splitedIdTag = idTag.toString().substring (0, 3);
                if(splitedIdTag == "Tab"){
                  tableCpationForTableAsImage(captionTextFrame, imageRectangle);
		  	}
		   }                        
        }        
        else if(imageStyle=="STYLE_3")
        {
            captionTextFrame.label="SideCaptionFrame";
            myDoc.align([imageRectangle.parent],AlignOptions.RIGHT_EDGES,AlignDistributeBounds.MARGIN_BOUNDS);
            if(myTemplateName == "Large_Template")
            {
                if(parseInt(imagePage.name)%2==0)
                    captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[0],right,imageRectangle.geometricBounds[2],right+39]
                else if(parseInt(imagePage.name)%2!=0)
                    captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[0],left,imageRectangle.geometricBounds[2],left+39]
            }
            else if(myTemplateName == "Medium_Template")
            {
                if(parseInt(imagePage.name)%2==0)
                    captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[0],right,imageRectangle.geometricBounds[2],right+34]
                else if(parseInt(imagePage.name)%2!=0)
                    captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[0],left,imageRectangle.geometricBounds[2],left+34]
            }
            else if(myTemplateName == "Small_Extended_Template")
            {
                if(parseInt(imagePage.name)%2==0)
                    captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[0],right,imageRectangle.geometricBounds[2],right+37]
                else if(parseInt(imagePage.name)%2!=0)
                    captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[0],left,imageRectangle.geometricBounds[2],left+37]
            }
            else if(myTemplateName == "Small_Condensed_Template")
            {
                if(parseInt(imagePage.name)%2==0)
                    captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[0],right,imageRectangle.geometricBounds[2],right+38]
                else if(parseInt(imagePage.name)%2!=0)
                    captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[0],left,imageRectangle.geometricBounds[2],left+38]
            }
            captionTextFrame.placeXML(captionNode);
            captionTextFrame.appliedObjectStyle = myDoc.objectStyles.item("FigSideCaption");
            
            if(captionTextFrame.overflows){
                captionTextFrame.remove();
                captionTextFrame=changeCaptionFrameToFullWidth(imageRectangle,captionNode);
                changeCaptionFrameObjectStyleForMultipleLines (captionTextFrame);
                var captionTag = captionTextFrame.associatedXMLElement;
                var parentTag = captionTag.parent.parent;
                if(parentTag.markupTag.name == "Figure"){
                    var idTag = parentTag.xmlAttributes.item("ID").value;
                    var splitedIdTag = idTag.toString().substring (0, 3);
                    if(splitedIdTag == "Tab"){
                      tableCpationForTableAsImage(captionTextFrame, imageRectangle);
                }
                }
            else{
                   captionTextFrame.insertionPoints[0].paragraphs[0].appliedParagraphStyle=myDoc.paragraphStyles.item("Figure_Caption_Side");
                }
            }
        }
        else if(imageStyle=="STYLE_4"){
            captionTextFrame.remove();
            captionTextFrame=changeCaptionFrameToFullWidth(imageRectangle,captionNode);
            changeCaptionFrameObjectStyleForMultipleLines (captionTextFrame);
            var captionTag = captionTextFrame.associatedXMLElement;
            var parentTag = captionTag.parent.parent;
            if(parentTag.markupTag.name == "Figure"){
                var idTag = parentTag.xmlAttributes.item("ID").value;
                var splitedIdTag = idTag.toString().substring (0, 3);
                if(splitedIdTag == "Tab"){
                  tableCpationForTableAsImage(captionTextFrame, imageRectangle);
		  	}
		   }                        
            }
        else if(imageStyle=="STYLE_5"){
            imageRectangle.parent.rotationAngle = 90;
            imageRectangle.parent.move([left, pageHeight-bottom]);
            var imageHeight = imageRectangle.geometricBounds[3] - imageRectangle.geometricBounds[1];
            var imageWidth = imageRectangle.geometricBounds[2] - imageRectangle.geometricBounds[0];
            
            var maxCaptionHeight = marginWidth - imageHeight;
            var moreWrap = marginHeight - imageWidth;
            captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[3],imageRectangle.geometricBounds[0],imageRectangle.geometricBounds[3]+maxCaptionHeight,imageRectangle.geometricBounds[0]+marginHeight]
            captionTextFrame.placeXML(captionNode);
            captionTextFrame.rotationAngle = 90;
            captionTextFrame.move([imageRectangle.geometricBounds[3], pageHeight-bottom])
            
            captionTextFrame.appliedObjectStyle = myDoc.objectStyles.item("FigureCaptionStyle");
            captionTextFrame.fit(FitOptions.frameToContent);            
            var imageCaptionGroup=imagePage.groups.add([imageRectangle.parent,captionTextFrame])
            imageRemainSpace = marginWidth - (imageCaptionGroup.geometricBounds[3] - imageCaptionGroup.geometricBounds[1]);
            imageCaptionGroup.move ([left+(imageRemainSpace/2), imageCaptionGroup.geometricBounds[0]]);
            refresh();
            imageCaptionGroup.ungroup();
            var captionTag = captionTextFrame.associatedXMLElement;
            var parentTag = captionTag.parent.parent;
            if(parentTag.markupTag.name == "Figure"){
                var idTag = parentTag.xmlAttributes.item("ID").value;
                var splitedIdTag = idTag.toString().substring (0, 3);
                if(splitedIdTag == "Tab"){
                  tableCpationForTableAsImage(captionTextFrame, imageRectangle);
		  	}
		   } 
       defImageFrmWrap = imageRectangle.parent.textWrapPreferences.textWrapOffset;
       imageRectangle.parent.textWrapPreferences.textWrapOffset = [defImageFrmWrap[0], defImageFrmWrap[3], defImageFrmWrap[2], moreWrap];
            }
        
        return captionTextFrame;
    }

     function changeCaptionFrameToSingleWidth(imageRectangle,captionNode)
    {
        var imagePage=imageRectangle.parentPage;
        var captionTextFrame=imagePage.textFrames.add();
       var left=imagePage.marginPreferences.left;
        var right=imagePage.marginPreferences.right;
        var top=imagePage.marginPreferences.top;
        var bottom=imagePage.marginPreferences.bottom;
        var pageHeight=myDoc.documentPreferences.pageHeight;
        var pageWidth=myDoc.documentPreferences.pageWidth;
                 var marginHeight=pageHeight-top-bottom;
        captionTextFrame.label="figureCaptionFrame";;

            var singleTextFrameWidth=getSingleTextFrameWidthFromPage(imagePage);
            if(parseInt(imagePage.name)%2==0)
                captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[2],right,imageRectangle.geometricBounds[2]+marginHeight,right+singleTextFrameWidth]
            if(parseInt(imagePage.name)%2!=0)
                captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[2],left,imageRectangle.geometricBounds[2]+marginHeight,left+singleTextFrameWidth]
            captionTextFrame.placeXML(captionNode);
            captionTextFrame.appliedObjectStyle = myDoc.objectStyles.item("FigureCaptionStyle_1Col");
            refresh();
            captionTextFrame.fit(FitOptions.frameToContent);
            myDoc.align([captionTextFrame,imageRectangle],AlignOptions.HORIZONTAL_CENTERS,AlignDistributeBounds.ITEM_BOUNDS);            
            imageRectangle.fit(FitOptions.frameToContent);
            var captionTag = captionTextFrame.associatedXMLElement;
            var parentTag = captionTag.parent.parent;
            if(parentTag.markupTag.name == "Figure"){
                var idTag = parentTag.xmlAttributes.item("ID").value;
                var splitedIdTag = idTag.toString().substring (0, 3);
                if(splitedIdTag == "Tab"){
                  tableCpationForTableAsImage(captionTextFrame, imageRectangle);
		  	}
		   }                        
        return captionTextFrame;
    }

    function typeOfImage(imageRectangle)
    {
        var imageType="";
        var imagePage=imageRectangle.parentPage;
        var imageWidth=imageRectangle.geometricBounds[3]-imageRectangle.geometricBounds[1];
        var pageWidth= myDoc.documentPreferences.pageWidth-imagePage.marginPreferences.left-imagePage.marginPreferences.right;
        var singleTextFrameWidth=getSingleTextFrameWidthFromPage (imagePage);
        
        if(Math.round(imageWidth)<=Math.round(singleTextFrameWidth) )
        {
            if(imagePage.marginPreferences.columnCount>1)
                imageType="S";
        }
        else{
            if(Math.round(imageWidth)>Math.round(pageWidth)*0.50&&Math.round(imageWidth)<Math.round(pageWidth)*0.70 && imageRectangle.absoluteRotationAngle == 0)
                imageType="T";
            else    
                imageType="F";
        }
        if(Math.round(imageWidth)>=Math.round(pageWidth))
            imageType="F";
        return imageType;
    }

    function getSingleTextFrameWidthFromPage(parentPage)
    {
        var singleTextFrameWidth=-1;
        var numberOfColumn=parentPage.marginPreferences.columnCount;
        var textFrameWidth=myDoc.documentPreferences.pageWidth-parentPage.marginPreferences.left-parentPage.marginPreferences.right;
        if(numberOfColumn>=2)
        {
            textFrameWidth=textFrameWidth-(parentPage.marginPreferences.columnGutter*(numberOfColumn-1));
            singleTextFrameWidth=textFrameWidth/parseInt(numberOfColumn);    
        }   
        else
            singleTextFrameWidth=textFrameWidth;
        return singleTextFrameWidth;
    }

    function placeObject(objectCaptionGroup,objectInternalRefNode)
    {
        var objectCaptionGroupPage=objectCaptionGroup.parentPage;
        var objectInternalRefAttribute=objectInternalRefNode.xmlAttributes.item("repoID");
        if(objectInternalRefAttribute.isValid)
        {
            var styleName="";
            var imageRepoId=objectInternalRefAttribute.value;
            var imageRepositionXmlElement=getNodesFromAttributeTypeAndValue("reposition","cs:repoID",imageRepoId.toString());
            if(imageRepositionXmlElement!=null)
            {
                var figureXmlElement=imageRepositionXmlElement.xmlElements[0];
                var figureIdAttribute=figureXmlElement.xmlAttributes.item("ID");
                var floatingType=floatingElementType(objectInternalRefNode);
                var styleIdAttribute=figureXmlElement.xmlAttributes.item("Style");
                if(styleIdAttribute.isValid&&floatingType.toLowerCase()=="table")
                    styleName=styleIdAttribute.value;
                if(figureIdAttribute.isValid)
                {
                    var figureId=figureIdAttribute.value;
                    var bottomPlacedObject=bottomPlacementObjectWithinSamePage(objectCaptionGroupPage);
                    var isSamePage=checkObjectReferenceIsOnSamePage(objectCaptionGroupPage,figureId);
                    var bottomPlacedObject=bottomPlacementObjectWithinSamePage(objectCaptionGroupPage);
                    var topPlacedObject=topPlacementObjectWithinSamePage(objectCaptionGroupPage);
                    var remainSpace=myDoc.documentPreferences.pageHeight-objectCaptionGroupPage.marginPreferences.top-objectCaptionGroupPage.marginPreferences.bottom;
                    if(topPlacedObject==null&&bottomPlacedObject!=null)
                        remainSpace=(bottomPlacedObject.geometricBounds[0])-8-objectCaptionGroupPage.marginPreferences.top;
                    else if(topPlacedObject!=null&&bottomPlacedObject==null)
                        remainSpace=myDoc.documentPreferences.pageHeight-(topPlacedObject.geometricBounds[2]+8)-objectCaptionGroupPage.marginPreferences.bottom;
                    else if(topPlacedObject!=null&&bottomPlacedObject!=null)
                        remainSpace=(bottomPlacedObject.geometricBounds[0]-8)-(topPlacedObject.geometricBounds[2]+8)
                    var currentObjectHeight=objectCaptionGroup.geometricBounds[2]-objectCaptionGroup.geometricBounds[0]
                    var imageType=typeOfImage(objectCaptionGroup)
                    if(imageType=="F")
                    {
                        var isSpecialRule=isSpecialFullWidthPlacementCondition(objectCaptionGroupPage);
                        if((Math.round(remainSpace)<Math.round(currentObjectHeight))||isSpecialRule)
                            isSamePage=false;
                    }
                    else if(imageType=="S")
                    {
                        var remainSpaceInSecondFarme=calculateRemainSpaceForSingleWidth (objectCaptionGroupPage, objectCaptionGroup);
                        if(Math.round(remainSpaceInSecondFarme)<Math.round(currentObjectHeight-1))
                            isSamePage=false;
                    }
                    var isOutOfPage=isFigureOutOfPage (objectCaptionGroupPage,objectCaptionGroup)
                    if(isOutOfPage&&imageType=="F")
                        isSamePage=false;
                    else if(isOutOfPage&&imageType=="S")
                        isSamePage=false;
                    if(isSamePage)
                    {
                        var isOnPreviousPage=checkCalloutOnPreviousPage(objectCaptionGroupPage,figureId)
                        if(topSingleWidth!=null&&imageType=="F")
                        {
                            if(topSingleWidth.parentPage.id==objectCaptionGroupPage.id)
                                isOnPreviousPage=false; 
                        }
                        var isReferencesPage=checkForReferencesPage(objectCaptionGroupPage);
                        if(!isOnPreviousPage && !isReferencesPage)
                        {
                            myDoc.align([objectCaptionGroup],AlignOptions.BOTTOM_EDGES,AlignDistributeBounds.MARGIN_BOUNDS);
                            if(floatingType=='table')
                            {
                                var table=null;
                                if(objectCaptionGroup instanceof TextFrame)
                                    table=objectCaptionGroup.tables[0];
                                else
                                {
                                    for(var i=0;i<objectCaptionGroup.allPageItems.length;i++)
                                    {
                                        if(objectCaptionGroup.allPageItems[i].label.toLowerCase()=="table")
                                            table=objectCaptionGroup.allPageItems[i].tables[0];
                                    }
                                }
                                if(table!=null)
                                {
                                    var frameHeight=  objectCaptionGroup.geometricBounds[2]-objectCaptionGroup.geometricBounds[0];
                                    if(objectCaptionGroup.absoluteRotationAngle==90)
                                        frameHeight=  objectCaptionGroup.geometricBounds[3]-objectCaptionGroup.geometricBounds[1];
                                    while((frameHeight-(table.parent.textFramePreferences.insetSpacing[2]+table.parent.textFramePreferences.insetSpacing[0]))>table.height+1)
                                    {
                                        table.parent.fit(FitOptions.frameToContent);
                                        myDoc.align([objectCaptionGroup],AlignOptions.BOTTOM_EDGES,AlignDistributeBounds.MARGIN_BOUNDS);
                                        frameHeight=  objectCaptionGroup.geometricBounds[2]-objectCaptionGroup.geometricBounds[0];
                                        if(objectCaptionGroup.absoluteRotationAngle==90)
                                            frameHeight=  objectCaptionGroup.geometricBounds[3]-objectCaptionGroup.geometricBounds[1];
                                    }    
                                }
                            }
                        }
                        var objectPlacedTextFrame=findPlacedTextFrame (objectCaptionGroup);
                        var imageType=typeOfImage (objectCaptionGroup)
                        if(imageType=="S")
                        {
                            if(bottomPlacedObject!=null)
                            {
                                var height=bottomPlacedObject.geometricBounds[0];
                                var imageHeight=objectCaptionGroup.geometricBounds[2]-objectCaptionGroup.geometricBounds[0];
                                var bottomImageType=typeOfImage (bottomPlacedObject);
                                if(bottomImageType=="F")
                                    objectCaptionGroup.move([objectCaptionGroup.geometricBounds[1],height-(imageHeight+8)]);
                                else if(bottomImageType=="S"&&Math.round(objectCaptionGroup.geometricBounds[1])==Math.round(bottomPlacedObject.geometricBounds[1]))    
                                    objectCaptionGroup.move([objectCaptionGroup.geometricBounds[1],height-(imageHeight+8)]);
                            }
                            var isSameTextFrame=checkFigureCalloutExistOnsameTextFrame(objectPlacedTextFrame, figureId.toString(), objectCaptionGroup)
                            if(bottomPlacedObject!=null)
                            { 
                                var objectType=typeOfImage (bottomPlacedObject);
                                if(objectType=="F")
                                    isSameTextFrame=false;
                            } 
                            if(isOutOfPage&&imageType=="S")
                                isSameTextFrame=false;
                            if(objectCaptionGroup.absoluteRotationAngle == 90 && imageType == "S")
                                isSameTextFrame=true;    
                            var topSignelWidthPlacedObject=getTopSingleWidthObject(objectCaptionGroupPage);
                            if(topSignelWidthPlacedObject!=null)
                            {
                                if(Math.round(topSignelWidthPlacedObject.geometricBounds[1])!=Math.round(objectCaptionGroup.geometricBounds[1]))
                                    isSameTextFrame=false;
                            }
                            if(!isSameTextFrame)
                            {
                                var objectTextFrame=findPlacedTextFrame(objectCaptionGroup);
                                if(objectTextFrame!=null)
                                {
                                    var nextFrame=objectTextFrame.nextTextFrame;
                                    if(nextFrame.parentPage.id!=objectCaptionGroup.parentPage.id)
                                    {
                                        var objectNextPage=nextPage(objectCaptionGroup.parentPage)
                                        if(objectNextPage.isValid)
                                        {
                                            objectCaptionGroupPage=objectNextPage;
                                            objectCaptionGroup.remove();
                                            var floatingType=floatingElementType(objectInternalRefNode);
                                            var newObjectCaptionGroup=null;
                                            if(floatingType.toLowerCase()=="figure")
                                                newObjectCaptionGroup=makeImage(objectInternalRefNode, objectCaptionGroupPage); //recursiveCall
                                            if(floatingType.toLowerCase()=="table")
                                                newObjectCaptionGroup=makeTable(objectInternalRefNode, objectCaptionGroupPage); //recursiveCall
                                            if(newObjectCaptionGroup!=null)    
                                                objectCaptionGroup=newObjectCaptionGroup;
                                            var continuationTextFrame=tableContinuation(objectCaptionGroup,styleName);
                                            if(continuationTextFrame!=null)
                                                objectCaptionGroup=continuationTextFrame;
                                            typeOfObjectPlaced(objectCaptionGroup)
                                            return objectCaptionGroup;    
                                        }
                                    }
                                }
                                myDoc.align([objectCaptionGroup],AlignOptions.TOP_EDGES,AlignDistributeBounds.MARGIN_BOUNDS);
                                myDoc.align([objectCaptionGroup],AlignOptions.RIGHT_EDGES,AlignDistributeBounds.MARGIN_BOUNDS);
                                var topObject = topPlacementObjectWithinSamePage(objectCaptionGroupPage);
                                var topObjectPlacedTextFrame=findPlacedTextFrame(topObject);
                                var currentObjectTextFrame=findPlacedTextFrame(objectCaptionGroup);
                                if(topObjectPlacedTextFrame!=null&&currentObjectTextFrame!=null)
                                {
                                    if(topObjectPlacedTextFrame.id!=currentObjectTextFrame.id)
                                        topObject=getTopFullWidthObject(objectCaptionGroupPage);                                    
                                }
                                if(topObject!=null)
                                    objectCaptionGroup.move([objectCaptionGroup.geometricBounds[1],topObject.geometricBounds[2]+8])
                                else if(topObject==null)
                                    objectCaptionGroup.move([objectCaptionGroup.geometricBounds[1],objectCaptionGroupPage.marginPreferences.top])
                                var floatingType=floatingElementType (objectInternalRefNode)
                                if(floatingType.toLowerCase()=="figure")
                                    fitCaptionFrame(objectCaptionGroup);
                                var isAfterCitation=checkCalloutIsAfterObjectPlaced(figureId.toString(), objectCaptionGroup);
                                if(isAfterCitation)
                                {
                                    myDoc.align([objectCaptionGroup],AlignOptions.BOTTOM_EDGES,AlignDistributeBounds.MARGIN_BOUNDS);
                                    var bottomFullWidthObject=getBottomFullWidthObject(objectCaptionGroupPage);
                                    if(bottomFullWidthObject!=null)
                                    {
                                        var height=bottomFullWidthObject.geometricBounds[0];
                                        var imageHeight=objectCaptionGroup.geometricBounds[2]-objectCaptionGroup.geometricBounds[0];
                                        objectCaptionGroup.move([objectCaptionGroup.geometricBounds[1],height-(imageHeight+8)])
                                    }   
                                    if(bottomPlacedObject!=null)
                                    {
                                        var bottomObjectPlacedTextFrame=findPlacedTextFrame(bottomPlacedObject);
                                        var currentObjectTextFrame=findPlacedTextFrame(objectCaptionGroup);
                                        if(bottomObjectPlacedTextFrame!=null&&currentObjectTextFrame!=null)
                                        {
                                            var imageType=typeOfImage (bottomPlacedObject);
                                            if(imageType=="S")
                                            {
                                                if(bottomObjectPlacedTextFrame.id!=currentObjectTextFrame.id)
                                                    bottomPlacedObject=null;
                                            }
                                            if(bottomFullWidthObject!=null)
                                            {
                                                if(bottomFullWidthObject.geometricBounds[0]<bottomPlacedObject.geometricBounds[0])
                                                    bottomPlacedObject=bottomFullWidthObject;
                                            }
                                        }
                                        var height=myDoc.documentPreferences.pageHeight-objectCaptionGroupPage.marginPreferences.bottom+8;
                                        if(bottomPlacedObject!=null)
                                        height=bottomPlacedObject.geometricBounds[0];
                                        var imageHeight=objectCaptionGroup.geometricBounds[2]-objectCaptionGroup.geometricBounds[0];
                                        objectCaptionGroup.move([objectCaptionGroup.geometricBounds[1],height-(imageHeight+8)])
                                    }
                                }
                            }
                          else if(isSameTextFrame)
                            {
                                $.writeln("call outs for this object is on other page: ");
                                if(isOnPreviousPage)
                                {
                                        var objectTextFrame=findPlacedTextFrame(objectCaptionGroup);
                                        if(objectTextFrame!=null )
                                        {
                                                if(imageType=="S")
                                                {
                                                    var myPage = objectTextFrame.parentPage;
                                                    for (var j = 0; myPage.pageItems.length>j;  j++) 
                                                    {
                                                      var mypg = myPage.pageItems[j]
                                                       if(mypg.label.toLowerCase() =="table" || mypg.label.toLowerCase() == "figure")
                                                         {                                                        
                                                                var currentObjectHeight = mypg.geometricBounds[2]-mypg.geometricBounds[0];
                                                                var textFrameHeight=myDoc.documentPreferences.pageHeight-myDoc.pages[2].marginPreferences.top-myDoc.pages[2].marginPreferences.bottom
                                                                var nextObjectHeight  = objectCaptionGroup.geometricBounds[2]-objectCaptionGroup.geometricBounds[0];
                                                                var documentHeight=myDoc.documentPreferences.pageHeight-myDoc.pages[1].marginPreferences.top-myDoc.pages[1].marginPreferences.bottom;
                                                                var remainSpace= parseFloat (textFrameHeight) - parseFloat (currentObjectHeight);
                                                                var columnWidthForTwoCol = parseInt(((((parseInt(myDoc.documentPreferences.pageWidth)-parseInt(page.marginPreferences.left))-parseInt(page.marginPreferences.right))-parseInt(page.marginPreferences.columnGutter))/2));
                                                                var objectWidth=  parseInt(mypg.geometricBounds[3])-parseInt(mypg.geometricBounds[1]);
                                                                 if(columnWidthForTwoCol>objectWidth || columnWidthForTwoCol==objectWidth)
                                                                 {
                                                                    if(remainSpace < nextObjectHeight )
                                                                    {
                                                                         if(imageType=="S" && objectCaptionGroup.rotationAngle == 0)
                                                                         {
                                                                            myDoc.align([objectCaptionGroup],AlignOptions.TOP_EDGES,AlignDistributeBounds.MARGIN_BOUNDS);
                                                                            myDoc.align([objectCaptionGroup],AlignOptions.RIGHT_EDGES,AlignDistributeBounds.MARGIN_BOUNDS);
                                                                        }
                                                                        else
                                                                        {
                                                                              var objectNextPage=nextPage(objectCaptionGroup.parentPage)
                                                                                if(objectNextPage.isValid)
                                                                                 {
                                                                                       objectCaptionGroupPage=objectNextPage;
                                                                                        objectCaptionGroup.remove();
                                                                                        var floatingType=floatingElementType(objectInternalRefNode);
                                                                                        var newObjectCaptionGroup=null;
                                                                                        if(floatingType.toLowerCase()=="figure")
                                                                                            newObjectCaptionGroup=makeImage(objectInternalRefNode, objectCaptionGroupPage); //recursiveCall
                                                                                        if(floatingType.toLowerCase()=="table")
                                                                                        newObjectCaptionGroup=makeTable(objectInternalRefNode, objectCaptionGroupPage); //recursiveCall
                                                                                        if(newObjectCaptionGroup!=null)    
                                                                                            objectCaptionGroup=newObjectCaptionGroup;
                                                                                        var continuationTextFrame=tableContinuation(objectCaptionGroup,styleName);
                                                                                        if(continuationTextFrame!=null)
                                                                                            objectCaptionGroup=continuationTextFrame;
                                                                                        typeOfObjectPlaced(objectCaptionGroup)
                                                                                       landscapeTableCenterAlignment(objectCaptionGroup, objectCaptionGroupPage, styleName);                                                                            
                                                                                        return objectCaptionGroup;                                                          
                                                                                }
                                                                         }
                                                                    }// End of Remain Space condition
                                                                }// End of columnWidthForTwoCol condition
                                                            else{
                                                                        var objectNextPage=nextPage(objectCaptionGroup.parentPage)
                                                                            if(objectNextPage.isValid)
                                                                             {
                                                                                   objectCaptionGroupPage=objectNextPage;
                                                                                    objectCaptionGroup.remove();
                                                                                    var floatingType=floatingElementType(objectInternalRefNode);
                                                                                    var newObjectCaptionGroup=null;
                                                                                    if(floatingType.toLowerCase()=="figure")
                                                                                        newObjectCaptionGroup=makeImage(objectInternalRefNode, objectCaptionGroupPage); //recursiveCall
                                                                                    if(floatingType.toLowerCase()=="table")
                                                                                    newObjectCaptionGroup=makeTable(objectInternalRefNode, objectCaptionGroupPage); //recursiveCall
                                                                                    if(newObjectCaptionGroup!=null)    
                                                                                        objectCaptionGroup=newObjectCaptionGroup;
                                                                                    var continuationTextFrame=tableContinuation(objectCaptionGroup,styleName);
                                                                                    if(continuationTextFrame!=null)
                                                                                        objectCaptionGroup=continuationTextFrame;
                                                                                    typeOfObjectPlaced(objectCaptionGroup)
                                                                                   landscapeTableCenterAlignment(objectCaptionGroup, objectCaptionGroupPage, styleName);                                                                            
                                                                                    return objectCaptionGroup;                                                          
                                                                            }
                                                                }
                                                              
                                                          }
                                                     }
                                            }// end of image type condition
                                        }//end of ObjectTextFrame condition
                                }// end of isOnPrevious page condition
                            }// End of Else if condition for isSameTextFrame
                            var objectGroupTextFrame=findPlacedTextFrame(objectCaptionGroup)
                            if(objectGroupTextFrame!=null)
                            {
                                var isBottomSingleWidthObject=checkForSingleWidthBottomPlacedObject (objectGroupTextFrame)
                                if(isBottomSingleWidthObject)
                                {
                                    myDoc.align([objectCaptionGroup],AlignOptions.BOTTOM_EDGES,AlignDistributeBounds.MARGIN_BOUNDS);
                                    var bottomFullWidthObject=getBottomFullWidthObject(objectCaptionGroupPage);
                                    if(bottomFullWidthObject!=null)
                                    {
                                        var height=bottomFullWidthObject.geometricBounds[0];
                                        var imageHeight=objectCaptionGroup.geometricBounds[2]-objectCaptionGroup.geometricBounds[0];
                                        objectCaptionGroup.move([objectCaptionGroup.geometricBounds[1],height-(imageHeight+8)])
                                    }   
                                    var height=objectCaptionGroup.geometricBounds[2]-objectCaptionGroup.geometricBounds[0];
                                    var allBottomPlacedObjects=allBottomPlacedObject(objectCaptionGroupPage,objectCaptionGroup)
                                    for(var b=0;b<allBottomPlacedObjects.length;b++)
                                    {
                                        var imageType=typeOfImage (allBottomPlacedObjects[b]);
                                        var oldObjectPlacedTextFrame=findPlacedTextFrame (allBottomPlacedObjects[b])
                                        var currentObjectPlacedTextFrame=findPlacedTextFrame (objectCaptionGroup)
                                        if(imageType=="S")
                                        {
                                            if(oldObjectPlacedTextFrame!=null&&currentObjectPlacedTextFrame!=null)
                                            {
                                                if(Math.round(oldObjectPlacedTextFrame.geometricBounds[1])==Math.round(currentObjectPlacedTextFrame.geometricBounds[1]))
                                                {
                                                    allBottomPlacedObjects[b].move([allBottomPlacedObjects[b].geometricBounds[1],allBottomPlacedObjects[b].geometricBounds[0]-(height+8)])
                                                    var imageType=typeOfImage (allBottomPlacedObjects[b]);
                                                    objectCaptionGroup=allBottomPlacedObjects[b];
                                                }    
                                            }
                                        }
                                    }
                                }
                            }
                        }
                       else if(imageType=="T")
                        {
                            var isOutOfPage=isFigureOutOfPage (objectCaptionGroupPage,objectCaptionGroup);
                            if(isOutOfPage)
                            {
                              var objectNextPage=nextPage(objectCaptionGroup.parentPage)
                                if(objectNextPage.isValid)
                                {
                                    objectCaptionGroupPage=objectNextPage;
                                    objectCaptionGroup.remove();
                                    var floatingType=floatingElementType(objectInternalRefNode);
                                    var newObjectCaptionGroup=null;
                                    if(floatingType.toLowerCase()=="figure")
                                        newObjectCaptionGroup=makeImage(objectInternalRefNode, objectCaptionGroupPage); //recursiveCall
                                    if(newObjectCaptionGroup!=null)    
                                        objectCaptionGroup=newObjectCaptionGroup;
                                    var continuationTextFrame=tableContinuation(objectCaptionGroup,styleName);
                                    if(continuationTextFrame!=null)
                                        objectCaptionGroup=continuationTextFrame;
                                    typeOfObjectPlaced(objectCaptionGroup)
                                    return objectCaptionGroup;    
                                }
                            }
                
                        }
                    else if(imageType=="F")
                        {
                            var allBottomPlacedObjects=allBottomPlacedObject(objectCaptionGroupPage,objectCaptionGroup)
                            if(!isOnPreviousPage)
                            {
                                var currentPlacedImage=objectCaptionGroup;
                                for(var b=0;b<allBottomPlacedObjects.length;b++)
                                {
                                    var height=currentPlacedImage.geometricBounds[2]-currentPlacedImage.geometricBounds[0];
                                    allBottomPlacedObjects[b].move([allBottomPlacedObjects[b].geometricBounds[1],allBottomPlacedObjects[b].geometricBounds[0]-(height+8)])
                                    var imageType=typeOfImage(allBottomPlacedObjects[b]);
                                    if(imageType=="F")
                                        objectCaptionGroup=allBottomPlacedObjects[b];
                                }
                            }// End of isPreviousPage if condition
                            else
                            {
                                var objectTextFrame=findPlacedTextFrame(objectCaptionGroup);
                                
                                if(objectTextFrame!=null )
                                {
                                    var myPage = objectTextFrame.parentPage;
                                      for (var j = 0; myPage.pageItems.length>j;  j++) 
                                       {
                                            var mypg = myPage.pageItems[j]
                                            if(mypg.label.toLowerCase() =="table" || mypg.label.toLowerCase() == "figure")
                                            {
                                                var currentObjectHeight = mypg.geometricBounds[2]-mypg.geometricBounds[0];
                                                var textFrameHeight=myDoc.documentPreferences.pageHeight-myDoc.pages[2].marginPreferences.top-myDoc.pages[2].marginPreferences.bottom
                                                var nextObjectHeight  = objectCaptionGroup.geometricBounds[2]-objectCaptionGroup.geometricBounds[0];
                                                var currentObjectWidth = mypg.geometricBounds[3] - mypg.geometricBounds[1];
                                                var nextObjectWidth  = objectCaptionGroup.geometricBounds[3]-objectCaptionGroup.geometricBounds[1];
                                                
                                                var remainSpace= parseFloat (textFrameHeight) - parseFloat (currentObjectHeight);
                                             if(remainSpace<nextObjectHeight)
                                             {
                                                 if(currentObjectWidth!=nextObjectWidth)
                                                 {      
                                                         var textFrameWidth=((((myDoc.documentPreferences.pageWidth-page.marginPreferences.left)-page.marginPreferences.right)-page.marginPreferences.columnGutter)/2);

                                                        if(imageType=="S")
                                                            {
                                                            myDoc.align([objectCaptionGroup],AlignOptions.TOP_EDGES,AlignDistributeBounds.MARGIN_BOUNDS);
                                                            myDoc.align([objectCaptionGroup],AlignOptions.RIGHT_EDGES,AlignDistributeBounds.MARGIN_BOUNDS);
                                                            }
                                                        else
                                                          {
                                                
                                                                        var objectNextPage=nextPage(objectCaptionGroup.parentPage)
                                                                        if(objectNextPage.isValid)
                                                                         {
                                                                               objectCaptionGroupPage=objectNextPage;
                                                                                objectCaptionGroup.remove();
                                                                                var floatingType=floatingElementType(objectInternalRefNode);
                                                                                var newObjectCaptionGroup=null;
                                                                                if(floatingType.toLowerCase()=="figure")
                                                                                    newObjectCaptionGroup=makeImage(objectInternalRefNode, objectCaptionGroupPage); //recursiveCall
                                                                                if(floatingType.toLowerCase()=="table")
                                                                                newObjectCaptionGroup=makeTable(objectInternalRefNode, objectCaptionGroupPage); //recursiveCall
                                                                                if(newObjectCaptionGroup!=null)    
                                                                                    objectCaptionGroup=newObjectCaptionGroup;
                                                                                var continuationTextFrame=tableContinuation(objectCaptionGroup,styleName);
                                                                                if(continuationTextFrame!=null)
                                                                                    objectCaptionGroup=continuationTextFrame;
                                                                               landscapeTableCenterAlignment(objectCaptionGroup, objectCaptionGroupPage, styleName);                                                                            
                                                                                typeOfObjectPlaced(objectCaptionGroup)
                                                                                return objectCaptionGroup;                                                          
                                                                        }
                                                                    
                                                              } // end of else part
                                                    }
                                                }//End of not equal to part
                                            }
                                        }//end of for loop
                                }
                             }
                        }
                    }
                    else
                    {
                        var objectNextPage=nextPage(objectCaptionGroupPage)
                        if(objectNextPage.isValid)
                        {
                            objectCaptionGroupPage=objectNextPage;
                            objectCaptionGroup.remove();
                            var floatingType=floatingElementType(objectInternalRefNode);
                            var newObjectCaptionGroup=null;
                            if(floatingType.toLowerCase()=="figure")
                                newObjectCaptionGroup=makeImage(objectInternalRefNode, objectCaptionGroupPage); //recursiveCall
                            if(floatingType.toLowerCase()=="table")
                                newObjectCaptionGroup=makeTable(objectInternalRefNode, objectCaptionGroupPage); //recursiveCall
                            if(newObjectCaptionGroup!=null)    
                                objectCaptionGroup=newObjectCaptionGroup;
                            var continuationTextFrame=tableContinuation(objectCaptionGroup,styleName);
                            if(continuationTextFrame!=null){
                                objectCaptionGroup=continuationTextFrame;
                              } 
                          
                          /*dont Use this code//
//~                           if(floatingType.toLowerCase()=="figure"){
//~                               typeOfPlaceedImage = typeOfImage(objectCaptionGroup.rectangles[0]);
//~                               if(figureId.substring (0, 3).toLowerCase()=="tab" && floatingType.toLowerCase()=="figure" && typeOfPlaceedImage=="F" || typeOfPlaceedImage=="S"){
//~                                   abc = objectCaptionGroup.rectangles[0].geometricBounds[0] -objectCaptionGroup.geometricBounds[0]
//~                                   objectCaptionGroup.move([objectCaptionGroup.geometricBounds[1], objectCaptionGroup.geometricBounds[0]+abc])
//~                                   
//~                                   }
//~                               }
*/
                           landscapeTableCenterAlignment(objectCaptionGroup, objectCaptionGroupPage, styleName);
                            typeOfObjectPlaced(objectCaptionGroup)
                            return objectCaptionGroup;    
                        }    
                    }
                }
            }//End of main If Condition
        }
        var continuationTextFrame=tableContinuation(objectCaptionGroup,styleName);
        if(continuationTextFrame!=null)
            objectCaptionGroup=continuationTextFrame;
            landscapeTableCenterAlignment(objectCaptionGroup, objectCaptionGroupPage, styleName);  
        typeOfObjectPlaced(objectCaptionGroup);
        return objectCaptionGroup;    
    }

    function checkForReferencesPage(objectPage)
    {
        var isReferencePage=false;
        var referenceNode=getNodesFromAttributeTypeAndValue("Bibliography","","")
        if(referenceNode!=null)
        {
            var referencePage=referenceNode.insertionPoints[0].parentTextFrames[0].parentPage;
            if(referencePage.isValid)
            {
                if(objectPage.id==referencePage.id)
                {
                    isReferencePage=true;
                }
            }
        }
        return isReferencePage;
    }
    function checkObjectReferenceIsOnSamePage(parentPage,tableOrFigureId)
    {
        var chapterBackmatterNode = null;
        if(TemplateName=="2_Col_Journals.indt"||TemplateName=="1_Col_Journals.indt")
            chapterBackmatterNode=getNodesFromAttributeTypeAndValue("ArticleBackmatter","","");
        var literatureParentPage=null;
        var onSamePage=true;
        var newPage;
        var imageInternalRef= getNodesFromAttributeTypeAndValue("InternalRef","RefID",tableOrFigureId.toString());
        if(imageInternalRef!=null)
        {
            newPage=imageInternalRef.insertionPoints[0].parentTextFrames[0].parentPage;  
            if(chapterBackmatterNode!=null)
            {
                var literatureParentPage=chapterBackmatterNode.insertionPoints[0].parentTextFrames[0].parentPage;
                if((parseInt(newPage.name)>parseInt(parentPage.name)))
                    onSamePage=false;
            }
        }
        return onSamePage;
    }

    function checkFigureCalloutExistOnsameTextFrame(frameSelected,objectId,figGp)
    {
        var onSameFrame=true;
        var internalReferenceNode= getNodesFromAttributeTypeAndValue("InternalRef", "RefID", objectId.toString()) //getNodes(myDoc,".//InternalRef[@RefID='" + figId+ "']");
        if(internalReferenceNode!=null)
        {
            var newInsertionPoint=internalReferenceNode.insertionPoints[0];  
            if(Math.round(newInsertionPoint.parentTextFrames[0].geometricBounds[1])!=Math.round(figGp.geometricBounds[1])&&Math.round(newInsertionPoint.parentTextFrames[0].geometricBounds[1])>Math.round(figGp.geometricBounds[1]))
            {
                if(figGp.parentPage.name==newInsertionPoint.parentTextFrames[0].parentPage.name)
                    onSameFrame=false;
            }
        }
        return onSameFrame;
    }

    function checkCalloutIsAfterObjectPlaced(objectId,objectCaptionGroup)
    {
        var isAfterCitation=false;
        var objectPlacedFrame=findPlacedTextFrame (objectCaptionGroup);
        var internalReferenceNode= getNodesFromAttributeTypeAndValue("InternalRef", "RefID", objectId.toString()) //getNodes(myDoc,".//InternalRef[@RefID='" + figId+ "']");
        var internalReferenceTextFrame=internalReferenceNode.insertionPoints[0].parentTextFrames[0];
        if(objectPlacedFrame.id==internalReferenceTextFrame.id)
        {
            var internalReferenceLine=internalReferenceNode.lines.firstItem();
            if(internalReferenceLine.isValid)
            {
                var citationYCoordinates=internalReferenceLine.baseline+1
                if(Math.round(citationYCoordinates)>Math.round(objectCaptionGroup.geometricBounds[2]))
                    isAfterCitation=true;
            }
        }
        return isAfterCitation;
    }

    function findPlacedTextFrame(image)
    {
        var matchTextFrame=null;
        if(image!=null)
        {
            var allTextFrames=image.parentPage.textFrames;
            var singleTextFrameWidth=getSingleTextFrameWidthFromPage(image.parentPage);
            var startCoordinate=image.parentPage.marginPreferences.top;
            var textFrameHeight=null;
            for(var i=0;i<allTextFrames.length;i++)
            {
                var widthOfObject=allTextFrames[i].geometricBounds[3]-allTextFrames[i].geometricBounds[1];
                if((Math.round(startCoordinate-1)<=Math.round(allTextFrames[i].geometricBounds[0]))&&(image.id!=allTextFrames[i].id)&&(Math.round(singleTextFrameWidth)==Math.round(widthOfObject))&&(allTextFrames[i].label==""|| allTextFrames[i].label.indexOf ("Main")==0)&&(Math.round(allTextFrames[i].geometricBounds[1])==Math.round(image.geometricBounds[1])||Math.floor(allTextFrames[i].geometricBounds[1])==Math.floor(image.geometricBounds[1])))
                    matchTextFrame=allTextFrames[i];
            }
        }
        return matchTextFrame;
    }

    function nextPage(parentPage)
    {
        var nextPageIndex=parseInt(parentPage.name);
        var nextPage=myDoc.pages[nextPageIndex];    
        return nextPage; 
    }

    function typeOfObjectPlaced(floatObject)
    {
        var floatObjectPlacedFrame=findPlacedTextFrame(floatObject);
        if(floatObjectPlacedFrame!=null)
        {
            var firstLine=floatObjectPlacedFrame.lines.firstItem();
            if(firstLine.isValid&&floatObjectPlacedFrame.contents!="")
            {
                var yCoordinate=firstLine.baseline+1;
                var imageType=typeOfImage(floatObject)
                if(yCoordinate<floatObject.geometricBounds[2])
                {
                    if(imageType=="S")
                        bottomSingleWidth=floatObject;
                    else 
                        bottomFullWidth=floatObject;
                }
                else if(yCoordinate>floatObject.geometricBounds[2])
                {
                    if(imageType=="S")
                        topSingleWidth=floatObject;
                    else 
                        topFullWidth=floatObject;
                }
            }
            else if(floatObjectPlacedFrame.contents=="")
            {
                var imageType=typeOfImage(floatObject)
                if(imageType=="S")
                    topSingleWidth=floatObject;
                else 
                    topFullWidth=floatObject;
            }
        }
        lastImagePlacedPage=floatObject.parentPage;
    }

    function allBottomPlacedObject(currentObjectPage,currentObjectGroup)
    {
        var allBottomPlacedObject=new Array();
        var allTextFrames=currentObjectPage.textFrames;
        for(var p=0;p<allTextFrames.length;p++)
        {
            var isFloatingObject=checkGroupContainsImageOrTable (allTextFrames[p])
            if(isFloatingObject && currentObjectGroup.id!=allTextFrames[p].id)
            {
                var floatObjectPlacedFrame=findPlacedTextFrame(allTextFrames[p]);
                if(floatObjectPlacedFrame!=null)
                {
                    var firstLine=floatObjectPlacedFrame.lines.firstItem();
                    if(firstLine.isValid && floatObjectPlacedFrame.contents!="")
                    {
                        var yCoordinate=firstLine.baseline+1;
                        var imageType=typeOfImage(allTextFrames[p]);
                        if(Math.round(yCoordinate)<Math.round(allTextFrames[p].geometricBounds[2]))
                            allBottomPlacedObject.push(allTextFrames[p]);
                    }
                }
            }
        }
        var allGroups=currentObjectPage.groups;
        for(var p=0;p<allGroups.length;p++)
        {
            var isFloatingObject=checkGroupContainsImageOrTable (allGroups[p])
            if(isFloatingObject && currentObjectGroup.id!=allGroups[p].id)
            {
                var floatObjectPlacedFrame=findPlacedTextFrame (allGroups[p]);
                if(floatObjectPlacedFrame!=null&&floatObjectPlacedFrame.contents!="")
                {
                    var firstLine=floatObjectPlacedFrame.lines.firstItem();
                    if(firstLine.isValid)
                    {
                        var yCoordinate=firstLine.baseline+1;
                        var imageType=typeOfImage(allGroups[p]);
                        if(Math.round(yCoordinate)<Math.round(allGroups[p].geometricBounds[2]))
                            allBottomPlacedObject.push(allGroups[p]);
                    }
                }
            }
        }
        return allBottomPlacedObject;
    }

    function topPlacementObjectWithinSamePage(currentFloatingObjectPage)
    {
        var lastObjectPlacedAtTop=null;
        var currentPage=currentFloatingObjectPage;
        var objectsInSamePage=getAllTopPlacedObjectInPage(currentPage)
        if(objectsInSamePage.length>0)
        {
            lastObjectPlacedAtTop=objectsInSamePage[0];
            for(var o=1;o<objectsInSamePage.length;o++)     //index should be started from 1 not from 0
            {
                if(Math.round (lastObjectPlacedAtTop.geometricBounds[2])<objectsInSamePage[o].geometricBounds[2])
                    lastObjectPlacedAtTop=objectsInSamePage[o];
            }
        }
        return lastObjectPlacedAtTop;
    }

    function bottomPlacementObjectWithinSamePage(currentFloatingObjectPage)
    {
        var lastObjectPlacedAtBottom=null;
        var currentPage=currentFloatingObjectPage;
        var objectsInSamePage=getAllBottomPlacedObjectInPage(currentPage)
        if(objectsInSamePage.length>0)
        {
            lastObjectPlacedAtBottom=objectsInSamePage[0];
            for(var o=1;o<objectsInSamePage.length;o++)     //index should be started from 1 not from 0
            {
                if(Math.round (lastObjectPlacedAtBottom.geometricBounds[2])<objectsInSamePage[o].geometricBounds[2])
                    lastObjectPlacedAtBottom=objectsInSamePage[o];
            }
        }
        return lastObjectPlacedAtBottom;
    }

    function getAllTopPlacedObjectInPage(currentFloatingObjectPage)
    {
        var topObjectsInSamePage=new Array();
        if(topFullWidth!=null)
        {
            if(topFullWidth.parentPage.id==currentFloatingObjectPage.id)
                topObjectsInSamePage.push(topFullWidth);
        }
        if(topSingleWidth!=null)
        {
            if(topSingleWidth.parentPage.id==currentFloatingObjectPage.id)
                topObjectsInSamePage.push(topSingleWidth);
        }
        return topObjectsInSamePage;
    }

    function getAllBottomPlacedObjectInPage(currentFloatingObjectPage)
    {
        var bottomObjectsInSamePage=new Array();
        if(bottomFullWidth!=null)
        {
            if(bottomFullWidth.parentPage.id==currentFloatingObjectPage.id)
                bottomObjectsInSamePage.push(bottomFullWidth);
        }
        if(bottomSingleWidth!=null)
        {
            if(bottomSingleWidth.parentPage.id==currentFloatingObjectPage.id)
                bottomObjectsInSamePage.push(bottomSingleWidth);
        }
        return bottomObjectsInSamePage;
    }

    function checkGroupContainsImageOrTable(group)
    {
        var isFloatingObject=false;
        if(group.constructor.name.toLowerCase()=="textframe") 
        {
            if(group.label.toLowerCase()=="table")
                isFloatingObject=true;
        }
        else if(group.constructor.name.toLowerCase()=="group")
        {
            var groupedTextFrames=group.textFrames;
            for(var gt=0;gt<groupedTextFrames.length;gt++)
            {
                if(groupedTextFrames[gt].label.toLowerCase()=="sidecaptionframe")
                isFloatingObject=true;
                else if(groupedTextFrames[gt].label.toLowerCase()=="figurecaptionframe")
                isFloatingObject=true;
                else if(groupedTextFrames[gt].label.toLowerCase()=="tablesidecaption")
                isFloatingObject=true;    
            }
            if(groupedTextFrames.length==0)
            {
                var allGrpahics=group.allGraphics;
                if(allGrpahics.length>0)
                isFloatingObject=true;
            }
        }
        return isFloatingObject;
    }

    function checkCalloutOnPreviousPage(objectPage,figureId)
    {
        var isOnPreviousPage=false;
        var imageInternalRef= getNodesFromAttributeTypeAndValue("InternalRef","RefID",figureId.toString());
        if(imageInternalRef!=null)
        {
            var imagePage=imageInternalRef.insertionPoints[0].parentTextFrames[0].parentPage;  
            if(parseInt(imagePage.name)<parseInt(objectPage.name))
                isOnPreviousPage=true; 
        }
        return isOnPreviousPage;
    }

    function calculateRemainSpaceForSingleWidth(objectCaptionGroupPage,objectCaptionGroup)
    {
        var remainSpace=0;
        var objectPlacedTextFrame=findPlacedTextFrame (objectCaptionGroup);
        var numberOfColumn=objectCaptionGroupPage.marginPreferences.columnCount;
        if(numberOfColumn==2&&objectPlacedTextFrame!=null)
        {
            var topObjectOnPage=null;
            var bottomObjectOnPage=null;
            var objectTextFrame=objectPlacedTextFrame.nextTextFrame;
            if(objectTextFrame.parentPage.id>objectPlacedTextFrame.parentPage.id)
            objectTextFrame=objectPlacedTextFrame;
            if(topSingleWidth!=null)
            {
                if(objectCaptionGroupPage.id==topSingleWidth.parentPage.id)
                {
                    var topSingleWidthPlacedTextFrame=findPlacedTextFrame (topSingleWidth);
                    if(topSingleWidthPlacedTextFrame.id==objectTextFrame.id)
                        topObjectOnPage=topSingleWidth;
                }
            }
            if(bottomSingleWidth!=null)
            {
                if(objectCaptionGroupPage.id==bottomSingleWidth.parentPage.id)
                {
                    var bottomSingleWidthPlacedTextFrame=findPlacedTextFrame (bottomSingleWidth);
                    if(bottomSingleWidthPlacedTextFrame.id==objectTextFrame.id)
                        bottomObjectOnPage=bottomSingleWidth;
                }
            }
            if(topObjectOnPage==null)
            {
                if(topFullWidth!=null)
                {
                    if(objectCaptionGroupPage.id==topFullWidth.parentPage.id)
                        topObjectOnPage=topFullWidth;
                }
            }
            if(bottomObjectOnPage==null)
            {
                if(bottomFullWidth!=null)
                {
                    if(objectCaptionGroupPage.id==bottomFullWidth.parentPage.id)
                        bottomObjectOnPage=bottomFullWidth;
                }
            }
            if(topObjectOnPage==null&&bottomObjectOnPage==null)
                remainSpace=myDoc.documentPreferences.pageHeight-objectCaptionGroupPage.marginPreferences.top-objectCaptionGroupPage.marginPreferences.bottom;
            else if(topObjectOnPage!=null&&bottomObjectOnPage==null)
                remainSpace=myDoc.documentPreferences.pageHeight-(topObjectOnPage.geometricBounds[2]+8)-objectCaptionGroupPage.marginPreferences.bottom;
            else if(topObjectOnPage==null&&bottomObjectOnPage!=null)
                remainSpace=myDoc.documentPreferences.pageHeight-objectCaptionGroupPage.marginPreferences.top-(bottomObjectOnPage.geometricBounds[2]-bottomObjectOnPage.geometricBounds[0])-8;
            else if(topObjectOnPage!=null&&bottomObjectOnPage!=null)
                remainSpace=(bottomObjectOnPage.geometricBounds[0]-8)-(topObjectOnPage.geometricBounds[2]+8);
        }
        return remainSpace;
    }

    function getTopSingleWidthObject(objectCaptionGroupPage)
    {
        var topSingleWidthObject=null;
        if(topSingleWidth!=null)
        {
            if(objectCaptionGroupPage.id==topSingleWidth.parentPage.id)
                topSingleWidthObject=topSingleWidth;
        }
        return topSingleWidthObject;
    }

    function getTopFullWidthObject(objectCaptionGroupPage)
    {
        var topFullWidthObject=null;
        if(topFullWidth!=null)
        {
            if(objectCaptionGroupPage.id==topFullWidth.parentPage.id)
                topFullWidthObject=topFullWidth;
        }
        return topFullWidthObject;
    }

    function getBottomFullWidthObject(objectCaptionGroupPage)
    {
        var bottomFullWidthObject=null;
        if(bottomFullWidth!=null)
        {
            if(objectCaptionGroupPage.id==bottomFullWidth.parentPage.id)
                bottomFullWidthObject=bottomFullWidth;
        }
        return bottomFullWidthObject;
    }

    function checkForSingleWidthBottomPlacedObject(objectTextFrame)
    {
        var isBottomSingleWidthObject=false;
        var currentFloatingObjectPage=objectTextFrame.parentPage;
        var allBottomPlacedObjects=getAllBottomPlacedObjectInPage(currentFloatingObjectPage)
        for(var b=0;b<allBottomPlacedObjects.length;b++)
        {
            var imageType=typeOfImage (allBottomPlacedObjects[b])
            if(imageType=="S")
            {
                var tempObjectTextFrame=findPlacedTextFrame (allBottomPlacedObjects[b])
                if(tempObjectTextFrame!=null)
                {
                    if(Math.round (tempObjectTextFrame.geometricBounds[1])==Math.round(objectTextFrame.geometricBounds[1]))
                        isBottomSingleWidthObject=true; 
                }
            }
        }
        return isBottomSingleWidthObject;
    }

    function changeCaptionFrameToFullWidth(imageRectangle,captionNode)
    {
        var imagePage=imageRectangle.parentPage;
        var captionTextFrame=imagePage.textFrames.add();
        var left=imagePage.marginPreferences.left;
        var right=imagePage.marginPreferences.right;
        var top=imagePage.marginPreferences.top;
        var bottom=imagePage.marginPreferences.bottom;
        var pageHeight=myDoc.documentPreferences.pageHeight;
        var pageWidth=myDoc.documentPreferences.pageWidth;
        var marginHeight=pageHeight-top-bottom;
        captionTextFrame.label="figureCaptionFrame"
        if(parseInt(imagePage.name)%2==0)
            captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[2],left,imageRectangle.geometricBounds[2]+marginHeight,pageWidth-left]
        if(parseInt(imagePage.name)%2!=0)
            captionTextFrame.geometricBounds=[imageRectangle.geometricBounds[2],left,imageRectangle.geometricBounds[2]+marginHeight,pageWidth-right]
        captionTextFrame.placeXML(captionNode);
        captionTextFrame.appliedObjectStyle = myDoc.objectStyles.item("FigureCaptionStyle");
        refresh();
        ctFbound = captionTextFrame.geometricBounds;
        captionTextFrame.fit(FitOptions.frameToContent);
        captionTextFrame.geometricBounds=[captionTextFrame.geometricBounds[0],captionTextFrame.geometricBounds[1],captionTextFrame.geometricBounds[2],captionTextFrame.geometricBounds[1]+(pageWidth-(right+left))]
        
        myDoc.align(imageRectangle,AlignOptions.LEFT_EDGES,AlignDistributeBounds.ITEM_BOUNDS)
        myDoc.align([imageRectangle.parent],AlignOptions.HORIZONTAL_CENTERS,AlignDistributeBounds.MARGIN_BOUNDS)
        return captionTextFrame;
    }

    function isSpecialFullWidthPlacementCondition(objectPlacementPage)
    {
        var isSpecialRule=false;
        var fullWidthObject=getBottomFullWidthObject(objectPlacementPage);
        var singleWidthObject=getTopSingleWidthObject(objectPlacementPage);
        var leftCoor=0;
        if(fullWidthObject!=null)
        {
            if(singleWidthObject!=null)
            {
                var singleWidthObjectPlacedFrame=findPlacedTextFrame (singleWidthObject)
                if(singleWidthObjectPlacedFrame!=null)
                {
                    if(parseInt(objectPlacementPage.name)%2==0)
                        leftCoor=objectPlacementPage.marginPreferences.left;
                    if(parseInt(objectPlacementPage.name)%2!=0)
                        leftCoor=objectPlacementPage.marginPreferences.right;
                    if(Math.round(leftCoor)!=Math.round(singleWidthObjectPlacedFrame.geometricBounds[1]))    
                        isSpecialRule=true;
                }
            }
        }
        return isSpecialRule;
    }

    function changeCaptionFrameObjectStyleForMultipleLines(captionFrame)
    {
        var isSingleLine=isCaptionSingleLiner(captionFrame)
        if(!isSingleLine){
            captionFrame.fit(FitOptions.frameToContent);             
            captionFrame.insertionPoints[0].paragraphs[0].appliedParagraphStyle=myDoc.paragraphStyles.item("Figure_Caption");
            captionFrame.appliedObjectStyle = myDoc.objectStyles.item("FigureTwoColumnCaptionStyle");
            }
        captionFrame.geometricBounds=[captionFrame.geometricBounds[0],captionFrame.geometricBounds[1],captionFrame.geometricBounds[2]+5,captionFrame.geometricBounds[3]]
        setSuperScriptSubscriptNoBreak(".//Subscript");        
        setSuperScriptSubscriptNoBreak(".//Superscript");
        setSuperScriptSubscriptNoBreak(".//cs_text[@type='superscript']");
        var textColumn=maxLineColumn(captionFrame);
        var val=textColumn.lines.lastItem().baseline +1;
        captionFrame.geometricBounds=[captionFrame.geometricBounds[0],captionFrame.geometricBounds[1],val,captionFrame.geometricBounds[3]];    
    }

    function isCaptionSingleLiner(captionFrame)
    {
        var isSingleLine=false;
        var numberOfLines=captionFrame.texts[0].lines.length;
        if(numberOfLines<=2)
            isSingleLine=true;
        return isSingleLine;
    }

    function maxLineColumn(boxTextFrame)
    {
        var textColumns=boxTextFrame.textColumns;
        var maxLinesColumn;
        var maxLines=0;
        for(i=0;i<textColumns.length;i++)
        {
            var lines=textColumns[i].lines.lastItem().baseline +textColumns[i].lines.lastItem().leading;
            if(maxLines<lines)
            {
                maxLinesColumn=textColumns[i];
                maxLines=lines;
            }
        }
        return maxLinesColumn;
    }
    
    function changeTableColumnSizeAttribute()
    {
        logMsg = logMsg + (++msgCounter) + ". Change Table Column Size Attribute. " + "\n"
        var xPath = "//table";
        var root  = myDoc.xmlElements[0];
        var node  = null;
        if(TemplateName=="3_Col_Journals.indt" || TemplateName=="2_Col_Journals.indt" || TemplateName=="1_Col_Journals.indt")
        {
            try{
                var xmlTableWidth=0;
                var xmlTableStyle="";
                var twoThirdCaptionWidth=0;
                var documentWidth=myDoc.documentPreferences.pageWidth-myDoc.pages[1].marginPreferences.left-myDoc.pages[1].marginPreferences.right;
                var documentHeight=myDoc.documentPreferences.pageHeight-myDoc.pages[1].marginPreferences.top-myDoc.pages[1].marginPreferences.bottom;
                var proc  = app.xmlRuleProcessors.add([xPath, ]);
                var match = proc.startProcessingRuleSet(root);
                var columnCount=myDoc.pages[1].marginPreferences.columnCount;
                if(columnCount==1)
                twoThirdCaptionWidth=41;
                else if(columnCount==2)
                twoThirdCaptionWidth=45;
                while( match!=undefined ) 
                { 
                    tableNode = match.element;
                    var tableAttributes=tableNode.xmlAttributes
                    var floatTypeAttribute=tableNode.xmlAttributes.itemByName("Float");
                    if(floatTypeAttribute.isValid)
                    {
                        floatType=floatTypeAttribute.value;
                    }
                    if(floatType.toLowerCase()=="yes")
                    {
                        for(var i=0;i<tableAttributes.length;i++)
                        {
                            if(tableAttributes[i].name.toLowerCase().indexOf("width")!=-1)
                                xmlTableWidth=tableAttributes[i].value;
                            if(tableAttributes[i].name.toLowerCase()=="style")
                                xmlTableStyle=tableAttributes[i].value;
                            if(tableAttributes[i].name.toLowerCase().indexOf("cs_col")!=-1)
                            {
                                var tableWidth=1;
                                if(xmlTableStyle.toLowerCase().indexOf("style4")!=-1)
                                    tableWidth=documentHeight;
                                else if(xmlTableStyle.toLowerCase().indexOf("style1")!=-1)
                                {
                                    if(myDoc.pages[1].marginPreferences.columnCount>1)
                                        tableWidth=(documentWidth-myDoc.pages[1].marginPreferences.columnGutter)/2;
                                    else
                                        tableWidth=documentWidth;
                                }
                                else if(xmlTableStyle.toLowerCase().indexOf("style2")!=-1)
                                    tableWidth=(documentWidth-twoThirdCaptionWidth);
                                else
                                    tableWidth=documentWidth;
                                if(parseInt(tableAttributes[i].value)==0)
                                {
                                    var tableColumnCount=getTableColumnCountFromXmlElement(tableNode);
                                    tableAttributes[i].value=""+parseFloat(tableWidth)/tableColumnCount+"";
                                }
                                else
                                    tableAttributes[i].value=""+parseFloat(tableWidth)*parseFloat(tableAttributes[i].value)/parseFloat(xmlTableWidth)+"";
                            }
                        }
                    }
                    match = proc.findNextMatch();
                }
            }catch( ex ){
                alert(ex);
            }
            finally{
                proc.endProcessingRuleSet();
                proc.remove();
            }
        }
    }

    function expandTable(tableElement,tablePage)
    {
        var table=tableElement.tables[0];
        var columnCnt=table.columns.length;
        for(var i = 0; i<columnCnt; i++)
        {
            var j=i+1;
            var col="cs_col"+j;
            var tableXMLElement=tableElement.parent.parent;
            var colwidthFromXml=tableXMLElement.xmlAttributes.item(col).value;
            var colwidth= colwidthFromXml;
            table.columns[i].width = colwidth;    
        }
    }

    function floatingElementType(objectInternalRefNode)
    {
        var floatingType=null;
        var objectInternalRefAttribute=objectInternalRefNode.xmlAttributes.item("repoID");
        if(objectInternalRefAttribute.isValid)
        {
            var floatingRepoId=objectInternalRefAttribute.value;
            var floatingRepositionXmlElement=getNodesFromAttributeTypeAndValue("reposition","cs:repoID",floatingRepoId.toString());
            if(floatingRepositionXmlElement!=null)
            {
                var floatingXmlElement=floatingRepositionXmlElement.xmlElements[0];
                floatingType=floatingXmlElement.markupTag.name;
            }
        }
        return floatingType;
    }

    function placeTwoThirdCaptionTableForJournals(tableFrame,isSingleColumn)
    {
        var tableGroup=null;
        var table=tableFrame.tables[0];
        if(table.rows.length>0)
        {
            var firstRowHeight=0;
            var tableGroupingItems=new Array();
            tableGroupingItems.push(tableFrame);
            var tableRows=table.rows;
            for(var tr=0;tr<tableRows.length;tr++)
            {
                var tableCells=tableRows[tr].cells;
                for(var tc=0;tc<tableCells.length;tc++)
                {
                    var relatedXmlElement=tableCells[tc].paragraphs[0].insertionPoints[0].associatedXMLElements;
                    var rowWidth=table.rows[0].height;
                    if(relatedXmlElement[0].markupTag.name.toLowerCase()=="cell")
                    {
                        var entryNodes=relatedXmlElement[0].xmlElements;
                        var len=entryNodes.length;
                        for(e=0;e<len;e++)
                        {
                            var entryNode=entryNodes[e];
                            if(entryNode.markupTag.name.toLowerCase()=="entry")
                            {
                                var entryElementParentNode=entryNode.xmlElements;
                                if(entryElementParentNode.length>0)
                                {
                                    for(var ep=0;ep<entryElementParentNode.length;ep++)
                                    {
                                        if(entryElementParentNode[ep].markupTag.name.toLowerCase()=="caption")
                                        {
                                            var tablePage=tableFrame.parentPage;
                                            var columnCount=tablePage.marginPreferences.columnCount;
                                            var tableCaptionFrame=tablePage.textFrames.add();
                                            try{
                                                tableCaptionFrame.appliedObjectStyle=myDoc.objectStyles.item("TableSideCaption");
                                            }catch(e){}
                                            tableCaptionFrame.label="TableSideCaption";
                                            tableCaptionFrame.geometricBounds=[tableFrame.geometricBounds[0],tableFrame.geometricBounds[1],tableFrame.geometricBounds[0]+20,table.width];
                                            tableCaptionFrame.placeXML(entryNode);
                                            tableCaptionFrame.fit(FitOptions.frameToContent);
                                            var obj = [tableCaptionFrame,tableFrame];
                                            var xCoordinates=0;
                                            if(isSingleColumn)
                                            {
                                                var tablePage=tableCaptionFrame.parentPage;
                                                if(parseInt(tablePage.name)%2==0)
                                                    xCoordinates=tablePage.marginPreferences.right;
                                                else
                                                    xCoordinates=tablePage.marginPreferences.left;
                                                tableCaptionFrame.move([xCoordinates,tableCaptionFrame.geometricBounds[0]])
                                            }
                                            else
                                                myDoc.distribute (obj,DistributeOptions.HORIZONTAL_CENTERS,AlignDistributeBounds.MARGIN_BOUNDS)
                                            tableCaptionFrame.geometricBounds=[tableCaptionFrame.geometricBounds[0],xCoordinates,tableCaptionFrame.geometricBounds[0]+20,tableFrame.geometricBounds[1]];
                                            tableCaptionFrame.fit(FitOptions.frameToContent);
                                            if(columnCount==1)    
                                                tableCaptionFrame.geometricBounds=[tableCaptionFrame.geometricBounds[0],xCoordinates,tableCaptionFrame.geometricBounds[0]+20,xCoordinates+39];
                                            else if(columnCount==2)
                                                tableCaptionFrame.geometricBounds=[tableCaptionFrame.geometricBounds[0],xCoordinates,tableCaptionFrame.geometricBounds[0]+20,xCoordinates+39];
                                            tableCaptionFrame.fit(FitOptions.frameToContent);    
                                            firstRowHeight=makeCellBlankAndReturnCellHeight (tableCells[tc]);
                                            tableCaptionFrame.move([tableCaptionFrame.geometricBounds[1],tableCaptionFrame.geometricBounds[0]+firstRowHeight])
                                            tableCaptionFrame.fit(FitOptions.frameToContent);
                                            tableFrame.fit(FitOptions.frameToContent);
                                            tableGroupingItems.push(tableCaptionFrame);
                                        }   
                                        if(entryElementParentNode[ep].markupTag.name.toLowerCase()=="tfooter")
                                        {
                                            var tablePage=tableFrame.parentPage;
                                            var columnCount=tablePage.marginPreferences.columnCount;
                                            var tableFooterFrame=tablePage.textFrames.add();
                                            try{
                                                tableFooterFrame.appliedObjectStyle=myDoc.objectStyles.item("TableSideFootnote");
                                            }catch(e){}
                                            tableFooterFrame.label="TableSideFooter";
                                            tableFooterFrame.geometricBounds=[tableFrame.geometricBounds[2]-20,tableFrame.geometricBounds[1],tableFrame.geometricBounds[2],table.width];
                                            tableFooterFrame.placeXML(entryNode);
                                            tableFooterFrame.fit(FitOptions.frameToContent);
                                            var obj = [tableFooterFrame,tableFrame];
                                            var xCoordinates=0;
                                            if(isSingleColumn)
                                            {
                                                var tablePage=tableFooterFrame.parentPage;
                                                if(parseInt(tablePage.name)%2==0)
                                                    xCoordinates=tablePage.marginPreferences.right;
                                                else
                                                    xCoordinates=tablePage.marginPreferences.left;
                                                tableFooterFrame.move([xCoordinates,tableFooterFrame.geometricBounds[0]])
                                            }
                                            else
                                                myDoc.distribute (obj,DistributeOptions.HORIZONTAL_CENTERS,AlignDistributeBounds.MARGIN_BOUNDS)
                                            tableFrame.fit(FitOptions.frameToContent);
                                            if(columnCount==1)    
                                                tableFooterFrame.geometricBounds=[tableFooterFrame.geometricBounds[0],tableFooterFrame.geometricBounds[1],tableFooterFrame.geometricBounds[0]+20,tableFooterFrame.geometricBounds[1]+39];
                                            else if(columnCount==2)    
                                                tableFooterFrame.geometricBounds=[tableFooterFrame.geometricBounds[0],tableFooterFrame.geometricBounds[1],tableFooterFrame.geometricBounds[0]+20,tableFooterFrame.geometricBounds[1]+39];
                                            tableFooterFrame.fit(FitOptions.frameToContent);
                                            tableFrame.fit(FitOptions.frameToContent);
                                            var cellHeight=makeCellBlankAndReturnCellHeight(tableCells[tc]);
                                            tableFrame.fit(FitOptions.frameToContent);
                                            var yTableCoordinates=tableFrame.geometricBounds[2]-(tableFooterFrame.geometricBounds[2]-tableFooterFrame.geometricBounds[0])-cellHeight;
                                            tableFooterFrame.move([tableFooterFrame.geometricBounds[1],yTableCoordinates])
                                            tableFooterFrame.fit(FitOptions.frameToContent);
                                            tableFrame.fit(FitOptions.frameToContent);
                                            tableGroupingItems.push(tableFooterFrame);
                                        }  
                                    }    
                                }
                            }
                        }
                    }
                }
            }
            if(tableGroupingItems.length>1)
                tableGroup=tablePage.groups.add(tableGroupingItems); 
            if(tableGroup!=null)
            {
                if(Math.round(tableGroup.geometricBounds[0])==Math.round(tableGroup.parentPage.marginPreferences.top))
                    tableGroup.move([tableGroup.geometricBounds[1],tableGroup.geometricBounds[0]-parseFloat(firstRowHeight)])
            }
        }
        return tableGroup;
    }

    function makeCellBlankAndReturnCellHeight(cellNode)
    {
        var cell=cellNode.insertionPoints[0].parent;
        try{
            cell.appliedCellStyle  = myDoc.cellStyles.item("Blank");
        }
        catch(e){}
        return cell.height;
    }

    function tableContinuation(objectPlaced,styleName)
    {
        myDoc.textPreferences.smartTextReflow = false;
        var table=null;
        if(objectPlaced instanceof TextFrame) 
        {
            var tables=objectPlaced.tables;
            if(tables.length>0)
                table=tables[0];
        }
        var continuationTextFrame=null;
        var isDetached=false;
        var continuedContents=null;
        while(objectPlaced.constructor.name.toLowerCase()=="textframe" &&objectPlaced.label.toLowerCase()=="table")
        {
            var objectParentPage=objectPlaced.parentPage;
            var nextTableFramePage=nextPage(objectParentPage);
              if(!nextTableFramePage.isValid){nextTableFramePage = myDoc.pages.add()}            
            var secondYCoordinate=objectPlaced.geometricBounds[2];
            var maxTableHeight=myDoc.documentPreferences.pageHeight-objectParentPage.marginPreferences.bottom;
            if(Math.round(secondYCoordinate)>Math.round(maxTableHeight))
                objectPlaced.geometricBounds=[objectPlaced.geometricBounds[0],objectPlaced.geometricBounds[1],maxTableHeight,objectPlaced.geometricBounds[3]];
            if(styleName.toLowerCase()=="style4")
            {
                if(nextTableFramePage.side==PageSideOptions.LEFT_HAND)
                    maxTableHeight=myDoc.documentPreferences.pageWidth-objectParentPage.marginPreferences.left;
                else if(nextTableFramePage.side==PageSideOptions.RIGHT_HAND)
                    maxTableHeight=myDoc.documentPreferences.pageWidth-objectParentPage.marginPreferences.right;
                secondYCoordinate=objectPlaced.geometricBounds[3];
                if(Math.round(secondYCoordinate)>Math.round(maxTableHeight))
                    objectPlaced.geometricBounds=[objectPlaced.geometricBounds[0],objectPlaced.geometricBounds[1],objectPlaced.geometricBounds[2],maxTableHeight+1];
            }
            if(objectPlaced.overflows)
            {
                if(styleName.toLowerCase()!="style4"){
                objectPlaced.geometricBounds = [objectPlaced.geometricBounds[0], objectPlaced.geometricBounds[1], objectPlaced.geometricBounds[2]+0.5, objectPlaced.geometricBounds[3]]                
                }
                else{
                objectPlaced.geometricBounds = [objectPlaced.geometricBounds[0], objectPlaced.geometricBounds[1], objectPlaced.geometricBounds[2], objectPlaced.geometricBounds[3]+0.5];                
                objectCaptionGroupPage = objectPlaced.parentPage;
                landscapeTableCenterAlignment(objectPlaced, objectCaptionGroupPage, styleName); 
                }
                var continuedContents=detachCaptionFromTable(objectPlaced,isDetached,table,continuedContents,styleName);
                if(continuedContents!=null)
                    isDetached=true;
                continuationTextFrame=nextTableFramePage.textFrames.add();
                continuationTextFrame.label="table";
                var xTableCoordinate=nextTableFramePage.marginPreferences.right
                var yTableCoordinate=nextTableFramePage.marginPreferences.right /***for width calculation ***/
                if(nextTableFramePage.side==PageSideOptions.LEFT_HAND)
                {
                    xTableCoordinate=nextTableFramePage.marginPreferences.right;
                    yTableCoordinate=nextTableFramePage.marginPreferences.left;
                }
                else if(nextTableFramePage.side==PageSideOptions.RIGHT_HAND)
                {
                    xTableCoordinate=nextTableFramePage.marginPreferences.left;
                    yTableCoordinate=nextTableFramePage.marginPreferences.right;
                }
                continuationTextFrame.geometricBounds=[nextTableFramePage.marginPreferences.top,xTableCoordinate,myDoc.documentPreferences.pageHeight-nextTableFramePage.marginPreferences.bottom,myDoc.documentPreferences.pageWidth-yTableCoordinate];
                if(styleName.toLowerCase()=="style4")
                {
                    var frameWidth=continuationTextFrame.geometricBounds[3]-continuationTextFrame.geometricBounds[1];
                    var frameHeight=continuationTextFrame.geometricBounds[2]-continuationTextFrame.geometricBounds[0];
                    continuationTextFrame.geometricBounds=[continuationTextFrame.geometricBounds[0],continuationTextFrame.geometricBounds[1],continuationTextFrame.geometricBounds[0]+frameWidth,continuationTextFrame.geometricBounds[1]+frameHeight+0.5];
                    myDoc.zeroPoint=[0,0];
                    continuationTextFrame.absoluteRotationAngle =90;
                    myDoc.zeroPoint=[0,0];
                    continuationTextFrame.move([continuationTextFrame.geometricBounds[1],myDoc.documentPreferences.pageHeight-continuationTextFrame.parentPage.marginPreferences.bottom])
                }
                
//~                 var captionTag = tableTextFrame.associatedXMLElement;
//~                 var parentTag = captionTag.parent;
//~                    if(!parentTag.xmlAttributes.item("ID").isValid){
                        continuationTextFrame.appliedObjectStyle = myDoc.objectStyles.item("FigureStyle");//Apply TableStyle Separate.
                   // }
                objectPlaced.nextTextFrame=continuationTextFrame;
                if(styleName.toLowerCase()!="style4")
                    continuationTextFrame.fit(FitOptions.frameToContent);
                objectPlaced=continuationTextFrame;
                objectPlaced.fit(FitOptions.FRAME_TO_CONTENT);
                var frameHeight = 0;
                var requireFrameHeight = 0;
                  if(styleName.toLowerCase()=="style4"){
                  frameHeight = objectPlaced.geometricBounds[3] - objectPlaced.geometricBounds[1];
                  requireFrameHeight =(myDoc.documentPreferences.pageWidth-myDoc.pages[1].marginPreferences.right) - 5;
                  }
                  else
                  {
                    frameHeight = objectPlaced.geometricBounds[2] - objectPlaced.geometricBounds[0];
                    requireFrameHeight =(myDoc.documentPreferences.pageHeight-myDoc.pages[1].marginPreferences.top) - 5;
                  }
                  if (frameHeight > requireFrameHeight)
                  {
                        if(styleName.toLowerCase()=="style4"){
                        var documentWidth=myDoc.documentPreferences.pageWidth-myDoc.pages[1].marginPreferences.left-myDoc.pages[1].marginPreferences.right;
                        var documentHeight=myDoc.documentPreferences.pageHeight-myDoc.pages[1].marginPreferences.top-myDoc.pages[1].marginPreferences.bottom;
                        objectPlaced.geometricBounds=[objectPlaced.geometricBounds[0],objectPlaced.geometricBounds[1],objectPlaced.geometricBounds[0]+(documentHeight),objectPlaced.geometricBounds[1]+(documentWidth)];
                        }                  
                  }
                  else
                  {
                objectPlaced.geometricBounds = [objectPlaced.geometricBounds[0], objectPlaced.geometricBounds[1], objectPlaced.geometricBounds[2], objectPlaced.geometricBounds[3]+0.5];                
                }
                
            }
            else
            {
                if(isDetached)
                    detachCaptionFromTable(objectPlaced, isDetached, table, continuedContents,styleName)
                if(continuedContents!=null&&continuationTextFrame!=null)
                    continuationTextFrame=continuationTextFrame.parent;
                break;
            }
        }
        myDoc.textPreferences.smartTextReflow = true;
        myDoc.save(SavePath);
        return continuationTextFrame;
    }

    function detachCaptionFromTable(objectPlaced,isDetached,table,continuedContents,styleName)
    {
        var objectPlacedPage=objectPlaced.parentPage;
        if(!isDetached)
        {
            var root=objectPlaced.associatedXMLElement;
            if(root.isValid)
            {
                var xPath = "//Caption"
                var xPath2="//CaptionNumber"
                var xPath3="//Caption/cs_text"
                var node  = null;
                var flag=true;
                var isCaptionNumberNode=false;
                var isColonNode=false;
                try{
                    var proc  = app.xmlRuleProcessors.add([xPath,xPath2,xPath3]);
                    var match = proc.startProcessingRuleSet(root);
                    while( match!=undefined ) 
                    {
                        node = match.element;
                        if(node.markupTag.name.toLowerCase()=="captionnumber" )
                        {
                            if(isColonNode)
                                continuedContents=node.contents+continuedContents;
                            else
                                continuedContents=node.contents;
                            isCaptionNumberNode=true;
                        }
                        else if(node.markupTag.name.toLowerCase()=="cs_text" &&isCaptionNumberNode)
                        {
                            if(node.contents.indexOf(":")!=-1)
                            {
                                continuedContents+=node.contents;
                                isColonNode=true;
                            }
                        }
                        else
                        {
                            var tableCaptionFrame=objectPlacedPage.textFrames.add();
                            var firstCell=node.insertionPoints[0].parent;
                            if(firstCell.constructor.name.toLowerCase()=="cell")
                            {
                                var firstCellHeight=firstCell.height;
                                var firstCellWidth=firstCell.width;
                                tableCaptionFrame.geometricBounds=[objectPlaced.geometricBounds[0],objectPlaced.geometricBounds[1],objectPlaced.geometricBounds[0]+firstCellHeight,objectPlaced.geometricBounds[1]+firstCellWidth];
                                tableCaptionFrame.textFramePreferences.ignoreWrap=true;
                                tableCaptionFrame.placeXML(node);
                                firstCell.height=firstCellHeight;
                            }
                            if(styleName.toLowerCase()=="style4")
                            {
                                myDoc.zeroPoint=[0,0];
                                tableCaptionFrame.absoluteRotationAngle =90;
                                myDoc.zeroPoint=[0,0];
                                tableCaptionFrame.geometricBounds=tableCaptionFrame.geometricBounds=[objectPlaced.geometricBounds[0],objectPlaced.geometricBounds[1],objectPlaced.geometricBounds[2],objectPlaced.geometricBounds[3]];
                                tableCaptionFrame.fit(FitOptions.frameToContent);
                            }
                        }
                        match = proc.findNextMatch(); 
                    }
                }catch( ex ){
                    alert(ex);
                }finally{
                    proc.endProcessingRuleSet();
                    proc.remove();
                }
            }
            isDetached=true;;
        }
        else
        {
            if(continuedContents!=null)
            {
                var tableCaptionFrame=objectPlacedPage.textFrames.add();
                if(styleName.toLowerCase()=="style4")
                    tableCaptionFrame.geometricBounds=[objectPlaced.geometricBounds[0],objectPlaced.geometricBounds[1],objectPlaced.geometricBounds[0]+4.2,objectPlaced.geometricBounds[1]+(objectPlaced.geometricBounds[2]-objectPlaced.geometricBounds[0])];
                else
                    tableCaptionFrame.geometricBounds=[objectPlaced.geometricBounds[0],objectPlaced.geometricBounds[1],objectPlaced.geometricBounds[0]+4.2,objectPlaced.geometricBounds[3]];
                tableCaptionFrame.contents=continuedContents;
                tableCaptionFrame.textFramePreferences.ignoreWrap=true;
                tableCaptionFrame.insertionPoints[-1].contents=SpecialCharacters.EN_SPACE;
                if(language.toLowerCase() =="en")
                    tableCaptionFrame.contents=tableCaptionFrame.contents+"(continued)";
                else if(language.toLowerCase() =="de")
                    tableCaptionFrame.contents=tableCaptionFrame.contents+"(Fortsetzung)";
// Akshay                                                                                           * * *updation for table continution for nl language * * *        
                 else if(language.toLowerCase() =="nl")
                 {
                 tableCaptionFrame.contents=tableCaptionFrame.contents+"(vervolg)";
                 tableCaptionFrame.applyObjectStyle(myDoc.objectStyles.item("TableCaptionContinue"), true);
                 tableCaptionFrame.textFramePreferences.ignoreWrap=false;
                 }
 //                                                                                                                             * * * * * 
                try{
                    tableCaptionFrame.insertionPoints[0].paragraphs[0].appliedParagraphStyle=myDoc.paragraphStyles.item("Table_Continued", true);
               
                }catch(e){}
                try{
                    tableCaptionFrame.insertionPoints[0].paragraphs[0].appliedParagraphStyle=myDoc.paragraphStyles.item("Table_Continued", true);
                }catch(e){}
                if(styleName.toLowerCase()=="style4")
                {
                    myDoc.zeroPoint=[0,0];
                    tableCaptionFrame.absoluteRotationAngle =90;
                    myDoc.zeroPoint=[0,0];
                    tableCaptionFrame.move([tableCaptionFrame.geometricBounds[1],myDoc.documentPreferences.pageHeight-objectPlacedPage.marginPreferences.bottom])
                }
                if(styleName.toLowerCase()=="style4")
                {
                    var newCellHeight=tableCaptionFrame.geometricBounds[3]-tableCaptionFrame.geometricBounds[1];
                    objectPlaced.move([objectPlaced.geometricBounds[1],objectPlaced.geometricBounds[2]]);            
                }
                else
                {
                    var newCellHeight=tableCaptionFrame.geometricBounds[2]-tableCaptionFrame.geometricBounds[0];
                    objectPlaced.move([objectPlaced.geometricBounds[1],objectPlaced.geometricBounds[0]]);            
                }
            }
            var objTableGroup = objectPlacedPage.groups.add([tableCaptionFrame, objectPlaced]);
            objTableGroup .label="table";
        }
        return continuedContents;
    }

    function isFigureOutOfPage(parentPage,objectPlaced)
    {
        var isFigureOutOfPage=false;
        var docHeight=myDoc.documentPreferences.pageHeight;
        var top=parentPage.marginPreferences.top;
        var bottom=parentPage.marginPreferences.bottom;
        var totalTextFrameHeight=docHeight-bottom;
        if(totalTextFrameHeight<objectPlaced.geometricBounds[2] && Math.floor(objectPlaced.geometricBounds[0])!=Math.floor(top))
            isFigureOutOfPage=true;
        return isFigureOutOfPage;
    }

    function fitCaptionFrame(objectPlaced)
    {
        var groupItems=objectPlaced.allPageItems;
        for(var i=0;i<groupItems.length;i++)
        {
            var groupItem=groupItems[i]
            if(groupItem instanceof TextFrame)
                groupItem.fit(FitOptions.FRAME_TO_CONTENT);
        }
    }

    function findWidthObjectPlacedInTextFrame(objectPlacedTextFrame)
    {
        var objectPlacedPage=objectPlacedTextFrame.parentPage;
        var allTextFrames=objectPlacedPage.textFrames;
        for(var t=0;t<allTextFrames.length;t++)
        { }
    }
    function BookMark()
    {
        logMsg = logMsg + (++msgCounter) + ". Create Bookmarks. " + "\n"
        var bookmark=null;
        var articleTitle=getNodesFromAttributeTypeAndValue ("frame/cs_text/ArticleTitle", "", "")
        if(articleTitle!=null)
        bookmark=addBookmarks (myDoc, articleTitle)
        if(bookmark!=null)
        {
            var sectionBookmark=bookmark;
            var previousSectionNumber=0;
            var headingNodes=getNodesForJournalsBookmarks("//Heading");
            var allSectionHeadingNodes=new Array();
            for(var i=0;i<headingNodes.length;i++)
            {
                try
                {
                    if(headingNodes[i].parent.markupTag.name.toLowerCase().indexOf("section")==0 )
                    {
                        var sectionNumber=headingNodes[i].parent.markupTag.name.toLowerCase().replace("section","");
                        if(sectionNumber>previousSectionNumber)
                            sectionBookmark=bookmark;
                        else if(sectionNumber<previousSectionNumber)
                            sectionBookmark=myDoc.bookmarks[0];
                        else if(sectionNumber==previousSectionNumber)
                            sectionBookmark=bookmark.parent;
                        bookmark=addBookmarks (sectionBookmark,headingNodes[i])
                        previousSectionNumber=sectionNumber;
                    }
                    if(headingNodes[i].parent.markupTag.name.toLowerCase()=="abstract"||headingNodes[i].parent.markupTag.name.toLowerCase().indexOf("bibliography")==0 )
                    {
                        if(myDoc.bookmarks.firstItem().isValid )
                            addBookmarks(myDoc.bookmarks.firstItem(),headingNodes[i])
                    }
                }catch(e){}
            }
        }    
    }

    function addBookmarks(rootBookmark,oneSectionNode)
    {
        var startingBookmark=null;
        var hyperlinkTextDestination=myDoc.hyperlinkTextDestinations.add(oneSectionNode.texts[0]);
        var noSpace = new RegExp("(\u0007)", "g");//|\ufeff=Xml elemnt \u2002=En_space
        var space = new RegExp("(\u2002|\u200A|\u2008|\u2009|\u2003)", "g");//|\ufeff=Xml elemnt \u2002=En_space
        var bookmarkContent=String(oneSectionNode.contents).replace(noSpace,"");
        bookmarkContent=String(bookmarkContent).replace(space," ");
        hyperlinkTextDestination.name=bookmarkContent;
        var startingBookmark=rootBookmark.bookmarks.add(hyperlinkTextDestination);
        return    startingBookmark;
    }

    function getNodesForJournalsBookmarks(expression)
    {
        var objectElement=new Array();
        var xPath = expression;
        var root  = myDoc.xmlElements[0];
        var node  = null;
        try{
            var proc  = app.xmlRuleProcessors.add([xPath, ]);
            var match = proc.startProcessingRuleSet(root);
            while( match!=undefined ) 
            { 
                node = match.element;
                var xmlAttribute=node.xmlAttributes.item("typename");
                objectElement.push(node)    
                match = proc.findNextMatch(); 
            }
        }catch( ex ){
            alert(ex);
        }
        finally{
            proc.endProcessingRuleSet();
            proc.remove();
        }
        return objectElement;
    }
    function changeRunningHeadTextForJournal(pageItem, runningHeadContent)
    {
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        app.findChangeGrepOptions.includeFootnotes = false;
        app.findChangeGrepOptions.includeHiddenLayers = false;
        app.findChangeGrepOptions.includeLockedLayersForFind = false;
        app.findChangeGrepOptions.includeLockedStoriesForFind = false;
        app.findChangeGrepOptions.includeMasterPages = true;
        app.findGrepPreferences = null;   
        app.findGrepPreferences.findWhat = "Running Head";
        var myFound = pageItem.findGrep();
        for (var i = 0; i < myFound.length; i++)
        {
            myFound[i].contents =runningHeadContent;
        }
        app.findGrepPreferences=NothingEnum.NOTHING;
        app.changeGrepPreferences=NothingEnum.NOTHING  
    }

    function mappingToLanguageName()
    {
        var indesignDocumentLanguage="";
        switch(language)
        {
            case "En":
                            indesignDocumentLanguage="English: USA Medical";
                            break;
            case "De":
                            indesignDocumentLanguage="German: 2006 Reform";
                            break;
            case "Nl":
                            indesignDocumentLanguage="Dutch: 2005 Reform";
                            break;                
        }
        return indesignDocumentLanguage;
    }
//========================================================================================
//  $Owner: Crest Premedia Solutions Pvt Ltd.
//  $Author: K.Velprakash.
//  $DateTime: 19/09/2013 10:50:03 AM$
//  $Software: InDesign CS5.5 JavaScript
//  $Version: #1 $
//  Copyright : Crest Premedia Solutions Pvt Ltd .All Rights Reserved.
//  $Purpose: Change the Language in all paragraph styles and character styles in Indesign File.
//========================================================================================
//oldLang = The Old Language which has to be change. Example:"German: 2006 Reform";
//newLang = The Old Language which has to be replace with. Example:"English: UK";
    function changeLanguage()
    {
        logMsg = logMsg + (++msgCounter) + ". Change Language. " + "\n"
        var newLang=mappingToLanguageName();
        var exceptionParaStyles=new Array();
        if(language.toLowerCase()=="de" || language.toLowerCase()=="nl"){    
            paraFile = File(sAppPath+"/"+"Scripts/Breeze/ingnoreParaStyle_DE.txt");
            }
        else{
            paraFile = File(sAppPath+"/"+"Scripts/Breeze/ingnoreParaStyle_EN.txt");            
            }
        paraFile.open ('r');
        while (!paraFile.eof){
            str= paraFile.readln();
            exceptionParaStyles.push(str)
            }
        paraFile.close();
        
        var paraStyles = myDoc.paragraphStyles;
        for(var i=0;i<paraStyles.length;i++)
        {
            var pStyle = paraStyles[i];
            if(pStyle.name!="[No Paragraph Style]")
            {
                if(!containsOf(exceptionParaStyles, pStyle.name))
                    pStyle.appliedLanguage = newLang.toString();
            }
        }
        var charStyles = myDoc.characterStyles;
        for(var j=0;j<charStyles.length;j++)
        {
            var cStyle = charStyles[j];
            if(cStyle.name!="[None]")
                cStyle.appliedLanguage =  newLang.toString();
        }
    }

    function containsOf(array,value)
    {
        var isHaving=false;
        for(var arr=0;arr<array.length;arr++)
        {
            if(array[arr].toLowerCase()==value.toLowerCase())
                isHaving=true;
        }
        return isHaving;
    }

    function columnReferenceBalance()
    {
        logMsg = logMsg + (++msgCounter) + ". Column Balance. " + "\n"
        try{
            var myDoc = app.documents[0];
            var parentEle = myDoc.xmlElements[0];
            var lastFrm = parentEle.xmlContent.insertionPoints[-1].parentTextFrames[0];
            lastFrm.textFramePreferences.verticalBalanceColumns = true;

            var myScriptPath = File (sAppPath+"/"+"Scripts/Breeze/ReferenceColumnBalance.jsxbin");
            app.doScript(myScriptPath);     
        }catch(e){logMsg = logMsg + (++msgCounter) + ". ERROR in Column Balance. " + "\n"+e+"\n"}
    }

    function logWritter(myPGXML,message) 
    {         
        aFile = File(myPGXML+"/"+ FolderName+"_Log.txt");
        var today = new Date();  
        if (!aFile.exists)
        {
            ErrorLog = zipPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
            NewErrorlog = ErrorLog.toString().split("/");
            ErrorLogTemp = NewErrorlog.pop();
            ErrorLogFileName = FolderName+"_Log.txt";
            aFile = File(ErrorLog+"/"+ErrorLogFileName);
        }
        aFile.open("w"); 
        aFile.write("   Crest PreMedia Solutions Pvt. Ltd.        "+"\n");
        aFile.write("   Dept : Technology (Java Team)        "+"\n");
        aFile.write("***********************************************"+"\n\n");
        aFile.write(message);
        aFile.write("\n\n"+"***********************************************"+"\n\n");
        var endDate = new Date();
        aFile.write("\n "+" Log creation finished on time : " + (endDate.toLocaleTimeString()) + " date : " + (endDate.toDateString()))
        aFile.close();  
    }


    function placeInlineObject()
    {
        logMsg = logMsg + (++msgCounter) + ". Placing Inline Objects. " + "\n"
        var   objectToSearch = ".//cs_repos[@repoID and @typename='figure']|.//cs_repos[@repoID and @typename='table'] ";
        var allTableAndFigureNode = getNodes(myDoc,objectToSearch);
        for(var i=0;i<allTableAndFigureNode.length;i++)
        {
            var objectNode=allTableAndFigureNode[i];
            var attributeName=objectNode.xmlAttributes.item("Type");
            if(objectNode.isValid)
            {
                if (objectNode.markupTag.name == "cs_repos") 
                {
                    if (objectNode.xmlAttributes.item("typename").value == "figure") 
                    {
                        placeInlineImage(objectNode);                        
                    }
                    if (objectNode.xmlAttributes.item("typename").value == "table") 
                    {
                        placeInlineTable(objectNode);
                    }
                }
            }
        }
        
    }

    function  placeInlineImage(objectNode)
    {
        var  nodeAttributeRepoid=objectNode.xmlAttributes.item('repoID').value;
        var singleFigureNodeWithRepoid=getNodes(myDoc, ".//cs:reposition[@cs:repoID="+nodeAttributeRepoid+"]/Figure");
        var myElement = singleFigureNodeWithRepoid[0];
        var checkFloat = myElement.xmlAttributes.item("Float").value;
        var figPath=myElement.xmlAttributes.item("cs_Fignode").value;
        if(checkFloat == "No")
        {
            var myDoc_ViewPref = myDoc.viewPreferences;
            myDoc_ViewPref.horizontalMeasurementUnits = MeasurementUnits.POINTS;
            myDoc_ViewPref. verticalMeasurementUnits = MeasurementUnits.POINTS;
            var rec=objectNode.insertionPoints[0].place(figPath);
            try
            {
                rec[0].parent.applyObjectStyle(myDoc.objectStyles.item("InlineImage"), true)
             }catch(ex)
            {}
            objectNode.xmlContent.lines[0].justification = Justification.CENTER_ALIGN;
            myDoc_ViewPref.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
            myDoc_ViewPref. verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;
        }
        
    }
    function placeInlineTable(objectNode)
    {
        var tableElement=null;
        var repoIdAttribute=objectNode.xmlAttributes.item("repoID");
        var repoId=repoIdAttribute.value;
        var tableElements=getNodes(myDoc, ".//cs:reposition[@cs:repoID='" + repoId + "']/table");
        if(tableElements!="")
        {
            tableElement=tableElements[0];
            var checkFloat = tableElement.xmlAttributes.item("Float").value;
             if(checkFloat == "No")
             {
                    var reqTableNode=tableElement.parent.xmlAttributes.item("cs:repoID").value;
                    var tableNode = getNodes(myDoc, ".//cs_repos[@repoID='" + reqTableNode + "']") ;
                    var nds=tableNode[0];
                    var paragraph=nds.insertionPoints[0].paragraphs[0];
                    var parentTextFrame=nds.insertionPoints[0].parentTextFrames[0];
                    var parentPage=nds.insertionPoints[0].parentTextFrames[0].parentPage;
                    var addBlankPage = parentPage;
                    var table  = tableElement.xmlElements[0].xmlElements[0].tables[0];
                    var rowCnt = table.columns.length;
                    var docWidth=myDoc.documentPreferences.pageWidth;
                    var docHeight=myDoc.documentPreferences.pageHeight;                 
                    var top=addBlankPage.marginPreferences.top;
                    var bottom=addBlankPage.marginPreferences.bottom;
                    var left=addBlankPage.marginPreferences.left;
                    var right=addBlankPage.marginPreferences.right;
                    var IsLandscape=false;
                    var tableStyle=tableElement.xmlAttributes.item("Style").value; 
                    try
                    {
                        tableElement.paragraphs[0].appliedParagraphStyle=myDoc.paragraphStyles.item("InlineTable");   
                    }catch(ex)
                    {alert(ex)}
                    for(var i = 0; i<rowCnt; i++)
                    {
                            var j=i+1;
                            var col="cs_col"+j;
                            //var textFrameWidth=myDoc.documentPreferences.pageWidth-left-right;
                            //var width=textFrameWidth;
                            //if(parentPage.marginPreferences.columnCount>1)
                              //  width=(textFrameWidth-(parentPage.marginPreferences.columnGutter*(parentPage.marginPreferences.columnCount-1)))/parentPage.marginPreferences.columnCount;
                            var colwidthFromXml=tableElement.xmlAttributes.item(col).value;
                            //var colwidth= (colwidthFromXml*width)/100;       
                            table.columns[i].width = colwidthFromXml;    
                    }
              }
             
        }    
    }

    function loadScript()
    {
        logMsg = logMsg + (++msgCounter) + ". Adjust author collection frame. " + "\n"
        var externalScriptsFolder=Folder(app.filePath+"/Scripts/Breeze/External_Lib/");
        
        var files=Folder(externalScriptsFolder).getFiles("*.*");
        for(var f=0;f<files.length;f++)
        {
            var fol=Folder(files[f])
            var jsxFiles=fol.getFiles ("*.jsx");
            for(var j=0;j<jsxFiles.length;j++)
            {
                app.doScript(jsxFiles[j]);
            }
        }
    }
    function findPage(theObj) {
         if (theObj.hasOwnProperty("baseline")) {
              theObj = theObj.parentTextFrames[0];
         }
         while (theObj != null) {
              if (theObj.hasOwnProperty ("parentPage")) return theObj.parentPage;
              var whatIsIt = theObj.constructor.name;
              switch (whatIsIt) {
                   case "Page" : return theObj;
                   case "Character" : theObj = theObj.parentTextFrames[0]; break;
                   case "Footnote" :; // drop through
                   case "Cell" : theObj = theObj.insertionPoints[0].parentTextFrames[0]; break;
                   case "Note" : theObj = theObj.storyOffset; break;
                   case "Application" : return null;
              }
              if (theObj == null) return null;
              theObj = theObj.parent;
         }
         return theObj
    } // end findPage

    function tocGeneration()
    {
        try
        {
            if(tocRequiredValue)
            {
                logMsg = logMsg + (++msgCounter) + ". TOC Creation. " + "\n"
                updateFirstNode();    
                myDoc.save(SavePath);
                var myScriptPath = File (sAppPath+"/"+"Scripts/Breeze/ElseVier_TOCGeneration.jsxbin");
                myScriptPath = myScriptPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
                Fpath=File(myScriptPath);
                myRefresh = Fpath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
                app.doScript(myRefresh);
            }
        }catch(e){
            logMsg = logMsg + (++msgCounter) + "Error in Toc Genaration. " + "\n";
        }    
    }


    function updateFirstNode()
    {
        var sectionsElement = getNodesFromAttributeTypeAndValue ("sections", "", "", myDoc.xmlElements[0])
        if(sectionsElement == null)
            sectionsElement  = myDoc.xmlElements[0];
        var sectionHeading = newGetNodes ("//section-title", sectionsElement);
        if(sectionHeading.length > 0 )
        {
            var paraStyle = sectionHeading[0].paragraphs[0].appliedParagraphStyle;
            if(paraStyle.isValid)
            {
                try{
                    refresh();
                }catch(e){
                }
            }
        }
    }

    function newGetNodes(expression, startElement)
    {
        var xmlElements=new Array();
        var xPath = expression;
        var root  = startElement;
        var node  = null;
        try{
            var proc  = app.xmlRuleProcessors.add([xPath, ]);
            var match = proc.startProcessingRuleSet(root);
            while( match!=undefined ) 
            { 
                node = match.element;
                xmlElements.push(node);
                match = proc.findNextMatch(); 
            }
        }
        catch( ex ){
            alert(ex);
        }
        finally{
            proc.endProcessingRuleSet();
            proc.remove();
        }
        return xmlElements;
    }
//Akshay crossMarkLogoPlacement                             * * * * *
function crossMarkLogoPlacement()
   {
    var crossMarkLayer = myDoc.layers.item("Cross Mark_Web Logo");
    var xPathToArticleInfo = ".//ArticleInfo";
    var articleType = getNodes(myDoc, xPathToArticleInfo);
    for (var i = 0; i < articleType.length; i++)
        {
          articleTypeTemp = articleType[i];
          var child = articleTypeTemp.xmlAttributes.item("ArticleType").value;
          var lowerName = toLowerCase(child);
          if(lowerName == "originalpaper" || lowerName = "bookreview" || lowerName = "briefcommunication" || lowerName = "continuingeducation" || lowerName = "editorialnotes" || lowerName = "interview" || child.name = "Letter" ||lowerName = "MediaReport" || lowerName = "ProductNotes" || lowerName = "Report" || lowerName = "ReviewPaper");
           {
            crossMarkLayer.visible  = true;
           }
        }
    
   }
function  toLowerCase(child)
{
    var lowerName = child.toLowerCase();
    return lowerName;
    }

// Cross Mark Logo Hyperlink and need DOI and System date format YYYY-MM-DD.
function crossMarkLogoPlacementHyperlink(){

        var doiChild = myDoc.xmlElements[0].xmlAttributes.item("cs_doi").value;    
        var logoPrefix = 'http://crossmark.crossref.org/dialog/?doi=';
        var logoSuffix = '&domain=pdf&date_stamp=';
        //var linkValue='10.1007/s11757-015-0322-7';
        var vDate = new Date();
        var vYear = vDate.getFullYear();
         var a =  vYear.toString();
        var vMonth = vDate.getMonth()+1;
        var b = vMonth.toString();
        var varDate = vDate.getDate();
        var c = varDate.toString();
        var finalDate = a + "-" + b + "-" + c;
      
        if(myTemplateName=="Vienna_40211_Template" || myTemplateName=="Vienna_501_Template" || myTemplateName=="Vienna_6Layout_Template"){
          var fPage = myDoc.masterSpreads.item("A-Text").rectangles;
          }
        else{
        var fPage = myDoc.pages.item(0).rectangles;
        }
        for(i = 0; i < fPage.length; i++){
            if(fPage[i].label=="CrossMark Logo"){
                
                var logoLinkValue = logoPrefix+doiChild+logoSuffix+finalDate;
                var myHyperlinkSource = myDoc.hyperlinkPageItemSources.add(fPage[i]);
                var logoHyperDest = myDoc.hyperlinkURLDestinations.add(logoLinkValue);
                var myHyperlink=myDoc.hyperlinks.add(myHyperlinkSource,logoHyperDest);
                //myHyperlink.visible;
                //myHyperlink.name = "myHyperlink1234"
            }
          }

}
// End of Cross Mark Logo.

 //                                                                                     * * * * * 
 
// White Space remove


    function removeWhiteSpaceEquation(){
        myEquationFiles = getEquationFiles();

        for(ab = 0; ab < myEquationFiles.length; ab++){
           replaceMathEquation(myEquationFiles[ab]); 
            }        
        }
    function replaceMathEquation(myFile){
        equationType = checkEquationType(myFile.name);
        tmpFile = myFile.parent.parent +"/" + myFile.name;
        tmpFileObj = File(tmpFile)
        tmpFileObj.open ('w')
        myFile.open('r')
        equationWithWhiteSpace = false;

        line = null;
        do{
            line = myFile.readln();
            if(line.substring (0, 14) == "%%BoundingBox:"){
                line_cont = line.split (" ")
                b1 = parseFloat (line_cont[1])+ 1.5
                if(equationType=="INLINE"){
                    b2 = parseFloat (line_cont[2])
                    }
                else{
                    b2 = parseFloat (line_cont[2])+ 1.5
                }
                b3 = parseFloat (line_cont[3])- 1.5
                if(equationType=="INLINE"){
                    b4 = parseFloat (line_cont[4])
                    }
                else{
                    b4 = parseFloat (line_cont[4])- 1.5
                }
                if(parseFloat (line_cont[1])==1.5){
                    equationWithWhiteSpace = true;
                    tmpFileObj.writeln (line)       
                    }
                else{
                    tmpFileObj.writeln ("%%BoundingBox: "+ b1 +" "+ b2 +" " + b3 +" "+ b4)       
                    }
                }
            else if(line.substring (0, 13) == "%%Dimensions:"){
                line_cont = line.split (" ")
                line_cont_H = line_cont[1].split ("=")
                line_cont_W = line_cont[2].split ("=")
                if(equationType=="INLINE"){
                    b1 = parseFloat (line_cont_H[1])
                    }
                else{
                    b1 = parseFloat (line_cont_H[1])- 0.0208
                }
                b2 = parseFloat (line_cont_W[1])- 0.0208
                if(equationWithWhiteSpace==false){
                    tmpFileObj.writeln ("%%Dimensions: H="+ b1 +", W="+ b2)       
                    }
                else{
                    tmpFileObj.writeln (line)       
                    }
                }
            else{
                    tmpFileObj.writeln (line)       
                }
            }
        while (!myFile.eof) {
        }
        myFile.close()

        tmpFileObj.copy (myFile)

        tmpFileObj.close()
        tmpFileObj.remove()
    }
    function getEquationFiles(){
        inddFilesName = myDoc.name.replace (".indd", "")
        epsFilesArray = new Array ()
        equationPath = myDoc.fullName.parent.parent +"/Graphics/Equation";
        epsFiles = Folder (equationPath).getFiles ("*.eps");
        return epsFiles;
        }
       function checkEquationType(myEqFileName){
        if(myEqFileName.indexOf (myDoc.name.replace(".indd", "")+"_Equ")!=-1){
            return "DISPLAY"
            }
        else{
            return "INLINE"        
            }
        }
        
        
            
        function landscapeTableCenterAlignment(objectCaptionGroup, objectCaptionGroupPage, styleName)
        {
              if(styleName == "Style4"){
      //center alignment                   
      myPage = objectCaptionGroupPage;
                   if( myTemplateName == "Large_Template" || myTemplateName == "Medium_Template"){
                         var documentWidth=myDoc.documentPreferences.pageWidth-myDoc.pages[1].marginPreferences.left-myDoc.pages[1].marginPreferences.right;
                        var documentHeight=myDoc.documentPreferences.pageHeight-myDoc.pages[1].marginPreferences.top-myDoc.pages[1].marginPreferences.bottom;
                        var columnWidthForTwoCol = ((((myDoc.documentPreferences.pageWidth-page.marginPreferences.left)-page.marginPreferences.right)-page.marginPreferences.columnGutter)/2);
                        var pageArea = myDoc.documentPreferences.pageWidth-myDoc.pages[1].marginPreferences.left-myDoc.pages[1].marginPreferences.right;
                       var tableHeight=  objectCaptionGroup.geometricBounds[2]-objectCaptionGroup.geometricBounds[0];
                       var tableWidth=  objectCaptionGroup.geometricBounds[3]-objectCaptionGroup.geometricBounds[1];
                        pageHeight=myDoc.documentPreferences.pageHeight;
                        pageWidth=myDoc.documentPreferences.pageWidth;
                        top=objectCaptionGroupPage.marginPreferences.top;
                        bottom=objectCaptionGroupPage.marginPreferences.bottom;
                      
                          if(myPage.marginPreferences.columnCount == 2)
                          {
                               if(columnWidthForTwoCol>tableWidth){                               
                                   available_space = (columnWidthForTwoCol-tableWidth)/2
                                            if(objectCaptionGroup.rotationAngle == 90){
                                                    objectCaptionGroup.move([objectCaptionGroup.geometricBounds[1]+available_space, pageHeight-bottom])
                                             }
                                         else{
                                                     objectCaptionGroup.move([objectCaptionGroup.geometricBounds[1]+available_space, top])
                                             }
                                   }         
                                  else{
                                        available_space = (pageArea-tableWidth)/2
                                           if(objectCaptionGroup.rotationAngle == 90){
                                                    objectCaptionGroup.move([objectCaptionGroup.geometricBounds[1]+available_space, pageHeight-bottom])
                                             }
                                         else{
                                                     objectCaptionGroup.move([objectCaptionGroup.geometricBounds[1]+available_space, top])
                                             }
                                    }                               
                           }                         
                      }//end of If condition of template checking     
                    if( myTemplateName == "Small_Condensed_Template" || myTemplateName == "Small_Extended_Template"){
                    pageHeight=myDoc.documentPreferences.pageHeight;
                    pageWidth=myDoc.documentPreferences.pageWidth;
                    bottom=objectCaptionGroupPage.marginPreferences.bottom;
                    top=objectCaptionGroupPage.marginPreferences.top;
                      var columnWidthSmall = myDoc.documentPreferences.pageWidth-myDoc.pages[1].marginPreferences.left-myDoc.pages[1].marginPreferences.right;
                      var tableWidthForSmall = objectCaptionGroup.geometricBounds[3]-objectCaptionGroup.geometricBounds[1];
                          if(columnWidthSmall>tableWidthForSmall){
                                      available_space = (columnWidthSmall-tableWidthForSmall)/2
                                      //objectCaptionGroup.move([objectCaptionGroup.geometricBounds[1]+available_space,pageHeight-bottom])
                                         if(objectCaptionGroup.rotationAngle == 90){
                                                    objectCaptionGroup.move([objectCaptionGroup.geometricBounds[1]+available_space, pageHeight-bottom])
                                             }
                                         else{
                                                     objectCaptionGroup.move([objectCaptionGroup.geometricBounds[1]+available_space, top])
                                             }
                              }                          
                      }//end of if condition of template small checking
                }//end of style4 condition
            }
            
        
function tableCpationForTableAsImage(captionTextFrame, imageRectangle){

                    var documentWidth = myDoc.documentPreferences.pageWidth;
                    var pageWidth = (documentWidth - (page.marginPreferences.left + page.marginPreferences.right))
                    var imageWidth = imageRectangle.geometricBounds[3]-imageRectangle.geometricBounds[1];
                    var numberOfColumn=myDoc.pages[0].marginPreferences.columnCount;
                    var columnWidthForTwoCol = ((((myDoc.documentPreferences.pageWidth-page.marginPreferences.left)-page.marginPreferences.right)-page.marginPreferences.columnGutter)/2);
                    var captionTag = captionTextFrame.associatedXMLElement;
                    var parentTag = captionTag.parent.parent;
                    if(parentTag.markupTag.name == "Figure"){
                        var idTag = parentTag.xmlAttributes.item("ID").value;
                        var splitedIdTag = idTag.toString().substring (0, 3);
                        if(splitedIdTag == "Tab"){
                            if(imageRectangle.parent.rotationAngle == 90){
                                captionStartPoint = captionTextFrame.geometricBounds[1];
                                imageStartPoint = imageRectangle.geometricBounds[1]
                                }
                            else{
                                captionStartPoint = captionTextFrame.geometricBounds[0];
                                imageStartPoint = imageRectangle.geometricBounds[0]
                                }
                            if(captionStartPoint > imageStartPoint){
                            imageRectangle.parent.appliedObjectStyle = myDoc.objectStyles.item("TableStyle");
                            if(imageRectangle.parent.rotationAngle == 90){
                                captionTextFrame.geometricBounds=[captionTextFrame.geometricBounds[0],captionTextFrame.geometricBounds[1]+5,captionTextFrame.geometricBounds[2],captionTextFrame.geometricBounds[3]+5]
                                
                            captionTextFrame.fit(FitOptions.FRAME_TO_CONTENT);                                        
                            captionTextFrame.move([imageRectangle.parent.geometricBounds[1], imageRectangle.parent.geometricBounds[2]]);
                            imageRectangle.parent.move([imageRectangle.parent.geometricBounds[1]+2.117+(captionTextFrame.geometricBounds[3]-captionTextFrame.geometricBounds[1]), imageRectangle.geometricBounds[2]]);
                                }
                            else{
                            captionTextFrame.geometricBounds = [captionTextFrame.geometricBounds[0]+5, captionTextFrame.geometricBounds[1], captionTextFrame.geometricBounds[2]+5, captionTextFrame.geometricBounds[3]];       
                            captionWidth = captionTextFrame.geometricBounds[3] - captionTextFrame.geometricBounds[1]
                            captionTextFrame.fit(FitOptions.FRAME_TO_CONTENT);          
                            captionTextFrame.geometricBounds = [captionTextFrame.geometricBounds[0], captionTextFrame.geometricBounds[1], captionTextFrame.geometricBounds[2], captionTextFrame.geometricBounds[1]+captionWidth]
                            captionTextFrame.move([captionTextFrame.geometricBounds[1], imageRectangle.geometricBounds[0]]);
                            imageRectangle.parent.move([imageRectangle.parent.geometricBounds[1], (imageRectangle.geometricBounds[0]+2.117+(captionTextFrame.geometricBounds[2]-captionTextFrame.geometricBounds[0]))]);
                                }                            
                            }
                                       imageRectangle.parent.appliedObjectStyle = myDoc.objectStyles.item("TableStyle");
                                if (numberOfColumn == 1){
                                        captionTextFrame.appliedObjectStyle = myDoc.objectStyles.item("TableCaptionStyle");
                                    }
                                else{
                                       if (captionTextFrame.lines.count() > 2){
                                           captionTextFrame.appliedObjectStyle = myDoc.objectStyles.item("TableTwoColumnCaptionStyle");
                                         }
                                     else{
                                         captionTextFrame.appliedObjectStyle = myDoc.objectStyles.item("TableCaptionStyle_1Col");
                                         }                                    
                                    }
                             }
                         }
}



function replaceOperatorSymbols(){
    var myOperatorSymbols = ["∕", "~", "÷", "∙", "≡", "≤", "≥", "+", "−", "±", ">", "<", "≈", "×", "≠", "="];
    for(i = 0; i < myOperatorSymbols.length; i++){
        findReplaceHyperlinkRegularFontEntity(myOperatorSymbols[i]);
        findReplaceRegularFontEntity(myOperatorSymbols[i]);
        findReplaceBoldFontEntity(myOperatorSymbols[i]);
        findReplaceItalicFontEntity(myOperatorSymbols[i]);
        findReplaceBoldItalicFontEntity(myOperatorSymbols[i]);
        
        findReplaceSuperScriptRegularFontEntity(myOperatorSymbols[i]);
        findReplaceSuperScriptBoldFontEntity(myOperatorSymbols[i]);
        findReplaceSuperScriptItalicFontEntity(myOperatorSymbols[i]);
        findReplaceSuperScriptBoldItalicFontEntity(myOperatorSymbols[i]);
        
        findReplaceSubScriptRegularFontEntity(myOperatorSymbols[i]);
        findReplaceSubScriptBoldFontEntity(myOperatorSymbols[i]);
        findReplaceSubScriptItalicFontEntity(myOperatorSymbols[i]);
        findReplaceSubScriptBoldItalicFontEntity(myOperatorSymbols[i]);        
        
        }
    }

function findReplaceHyperlinkRegularFontEntity(mySmbol){
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.includeFootnotes = true;
    app.findChangeTextOptions.includeHiddenLayers = false;
    app.findChangeTextOptions.includeLockedLayersForFind = false;
    app.findChangeTextOptions.includeLockedStoriesForFind = false;
    app.findChangeTextOptions.includeMasterPages = false;
    app.findChangeTextOptions.wholeWord = false;

    app.findTextPreferences.fontStyle = "Regular";
    app.findTextPreferences.fillColor = myDoc.swatches.item("Hyperlink");
    app.findTextPreferences.position = Position.NORMAL;
    app.findTextPreferences.findWhat = mySmbol;

    app.changeTextPreferences.appliedCharacterStyle = myDoc.characterStyles.item("cSYMBOL_Hyperlink");

    myDoc.changeText();

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;    
    }


function findReplaceRegularFontEntity(mySmbol){
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.includeFootnotes = true;
    app.findChangeTextOptions.includeHiddenLayers = false;
    app.findChangeTextOptions.includeLockedLayersForFind = false;
    app.findChangeTextOptions.includeLockedStoriesForFind = false;
    app.findChangeTextOptions.includeMasterPages = false;
    app.findChangeTextOptions.wholeWord = false;

    app.findTextPreferences.fontStyle = "Regular"
    app.findTextPreferences.baselineShift = 0;
    app.findTextPreferences.position = Position.NORMAL;
    app.findTextPreferences.findWhat = mySmbol;

    app.changeTextPreferences.appliedCharacterStyle = myDoc.characterStyles.item("cSYMBOL");

    myDoc.changeText();

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;    
    }

function findReplaceBoldFontEntity(mySmbol){
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.includeFootnotes = true;
    app.findChangeTextOptions.includeHiddenLayers = false;
    app.findChangeTextOptions.includeLockedLayersForFind = false;
    app.findChangeTextOptions.includeLockedStoriesForFind = false;
    app.findChangeTextOptions.includeMasterPages = false;
    app.findChangeTextOptions.wholeWord = false;

    app.findTextPreferences.fontStyle = "Bold"
    app.findTextPreferences.position = Position.NORMAL;
    app.findTextPreferences.findWhat = mySmbol;

    app.changeTextPreferences.appliedCharacterStyle = myDoc.characterStyles.item("cSYMBOL_Bold");

    myDoc.changeText();

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;    
    }

function findReplaceItalicFontEntity(mySmbol){
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.includeFootnotes = true;
    app.findChangeTextOptions.includeHiddenLayers = false;
    app.findChangeTextOptions.includeLockedLayersForFind = false;
    app.findChangeTextOptions.includeLockedStoriesForFind = false;
    app.findChangeTextOptions.includeMasterPages = false;
    app.findChangeTextOptions.wholeWord = false;

    app.findTextPreferences.fontStyle = "Italic"
    app.findTextPreferences.position = Position.NORMAL;
    app.findTextPreferences.findWhat = mySmbol;

    app.changeTextPreferences.appliedCharacterStyle = myDoc.characterStyles.item("cSYMBOL_Italic");

    myDoc.changeText();

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;    
    }

function findReplaceBoldItalicFontEntity(mySmbol){
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.includeFootnotes = true;
    app.findChangeTextOptions.includeHiddenLayers = false;
    app.findChangeTextOptions.includeLockedLayersForFind = false;
    app.findChangeTextOptions.includeLockedStoriesForFind = false;
    app.findChangeTextOptions.includeMasterPages = false;
    app.findChangeTextOptions.wholeWord = false;

    app.findTextPreferences.fontStyle = "Bold Italic"
    app.findTextPreferences.position = Position.NORMAL;
    app.findTextPreferences.findWhat = mySmbol;

    app.changeTextPreferences.appliedCharacterStyle = myDoc.characterStyles.item("cSYMBOL_BoldItalic");

    myDoc.changeText();

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;    
    }

function findReplaceSuperScriptRegularFontEntity(mySmbol){
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.includeFootnotes = true;
    app.findChangeTextOptions.includeHiddenLayers = false;
    app.findChangeTextOptions.includeLockedLayersForFind = false;
    app.findChangeTextOptions.includeLockedStoriesForFind = false;
    app.findChangeTextOptions.includeMasterPages = false;
    app.findChangeTextOptions.wholeWord = false;

    app.findTextPreferences.fontStyle = "Regular"
    app.findTextPreferences.position = Position.SUPERSCRIPT;
    app.findTextPreferences.findWhat = mySmbol;

    app.changeTextPreferences.appliedCharacterStyle = myDoc.characterStyles.item("cSYMBOL_Superscript");

    myDoc.changeText();

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;    
    }


function findReplaceSuperScriptBoldFontEntity(mySmbol){
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.includeFootnotes = true;
    app.findChangeTextOptions.includeHiddenLayers = false;
    app.findChangeTextOptions.includeLockedLayersForFind = false;
    app.findChangeTextOptions.includeLockedStoriesForFind = false;
    app.findChangeTextOptions.includeMasterPages = false;
    app.findChangeTextOptions.wholeWord = false;

    app.findTextPreferences.fontStyle = "Bold"
    app.findTextPreferences.position = Position.SUPERSCRIPT;
    app.findTextPreferences.findWhat = mySmbol;

    app.changeTextPreferences.appliedCharacterStyle = myDoc.characterStyles.item("cSYMBOL_Superscript_Bold");

    myDoc.changeText();

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;    
    }


function findReplaceSuperScriptItalicFontEntity(mySmbol){
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.includeFootnotes = true;
    app.findChangeTextOptions.includeHiddenLayers = false;
    app.findChangeTextOptions.includeLockedLayersForFind = false;
    app.findChangeTextOptions.includeLockedStoriesForFind = false;
    app.findChangeTextOptions.includeMasterPages = false;
    app.findChangeTextOptions.wholeWord = false;

    app.findTextPreferences.fontStyle = "Italic"
    app.findTextPreferences.position = Position.SUPERSCRIPT;
    app.findTextPreferences.findWhat = mySmbol;

    app.changeTextPreferences.appliedCharacterStyle = myDoc.characterStyles.item("cSYMBOL_Superscript_Italic");

    myDoc.changeText();

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;    
    }


function findReplaceSuperScriptBoldItalicFontEntity(mySmbol){
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.includeFootnotes = true;
    app.findChangeTextOptions.includeHiddenLayers = false;
    app.findChangeTextOptions.includeLockedLayersForFind = false;
    app.findChangeTextOptions.includeLockedStoriesForFind = false;
    app.findChangeTextOptions.includeMasterPages = false;
    app.findChangeTextOptions.wholeWord = false;

    app.findTextPreferences.fontStyle = "Bold Italic"
    app.findTextPreferences.position = Position.SUPERSCRIPT;
    app.findTextPreferences.findWhat = mySmbol;

    app.changeTextPreferences.appliedCharacterStyle = myDoc.characterStyles.item("cSYMBOL_Superscript_Bold_Italic");

    myDoc.changeText();

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;    
    }





function findReplaceSubScriptRegularFontEntity(mySmbol){
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.includeFootnotes = true;
    app.findChangeTextOptions.includeHiddenLayers = false;
    app.findChangeTextOptions.includeLockedLayersForFind = false;
    app.findChangeTextOptions.includeLockedStoriesForFind = false;
    app.findChangeTextOptions.includeMasterPages = false;
    app.findChangeTextOptions.wholeWord = false;

    app.findTextPreferences.fontStyle = "Regular"
    app.findTextPreferences.position = Position.SUBSCRIPT;
    app.findTextPreferences.findWhat = mySmbol;

    app.changeTextPreferences.appliedCharacterStyle = myDoc.characterStyles.item("cSYMBOL_Subscript");

    myDoc.changeText();

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;    
    }


function findReplaceSubScriptBoldFontEntity(mySmbol){
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.includeFootnotes = true;
    app.findChangeTextOptions.includeHiddenLayers = false;
    app.findChangeTextOptions.includeLockedLayersForFind = false;
    app.findChangeTextOptions.includeLockedStoriesForFind = false;
    app.findChangeTextOptions.includeMasterPages = false;
    app.findChangeTextOptions.wholeWord = false;

    app.findTextPreferences.fontStyle = "Bold"
    app.findTextPreferences.position = Position.SUBSCRIPT;
    app.findTextPreferences.findWhat = mySmbol;

    app.changeTextPreferences.appliedCharacterStyle = myDoc.characterStyles.item("cSYMBOL_Subscript_Bold");

    myDoc.changeText();

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;    
    }


function findReplaceSubScriptItalicFontEntity(mySmbol){
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.includeFootnotes = true;
    app.findChangeTextOptions.includeHiddenLayers = false;
    app.findChangeTextOptions.includeLockedLayersForFind = false;
    app.findChangeTextOptions.includeLockedStoriesForFind = false;
    app.findChangeTextOptions.includeMasterPages = false;
    app.findChangeTextOptions.wholeWord = false;

    app.findTextPreferences.fontStyle = "Italic"
    app.findTextPreferences.position = Position.SUBSCRIPT;
    app.findTextPreferences.findWhat = mySmbol;

    app.changeTextPreferences.appliedCharacterStyle = myDoc.characterStyles.item("cSYMBOL_Subscript_Italic");

    myDoc.changeText();

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;    
    }


function findReplaceSubScriptBoldItalicFontEntity(mySmbol){
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.includeFootnotes = true;
    app.findChangeTextOptions.includeHiddenLayers = false;
    app.findChangeTextOptions.includeLockedLayersForFind = false;
    app.findChangeTextOptions.includeLockedStoriesForFind = false;
    app.findChangeTextOptions.includeMasterPages = false;
    app.findChangeTextOptions.wholeWord = false;

    app.findTextPreferences.fontStyle = "Bold Italic";
    app.findTextPreferences.position = Position.SUBSCRIPT;
    app.findTextPreferences.findWhat = mySmbol;

    app.changeTextPreferences.appliedCharacterStyle = myDoc.characterStyles.item("cSYMBOL_Subscript_Bold_Italic");

    myDoc.changeText();

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;    
    }

function FindChangeByList(){
        
          var myScriptPath = File (sAppPath+"/"+"Scripts/Breeze/FindChangeByList.jsxbin");
        myScriptPath = myScriptPath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        Fpath=File(myScriptPath);
        myRefresh = Fpath.toString().replace(/\%20/g," ").replace(/\/([a-z])/g,"$1:");
        app.doScript(myRefresh);
    
    }