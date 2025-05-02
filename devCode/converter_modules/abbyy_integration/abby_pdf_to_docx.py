import win32com.client as win32c
from devCode.converter_modules.abbyy_integration import SamplesConfig
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(filename)s - %(message)s')

def Run(file_path,output_path):
    ## Load ABBYY FineReader Engine
    LoadEngine()
    try:
        ## Process with ABBYY FineReader Engine
        ProcessWithEngine(file_path,output_path)
    finally:
	## Unload ABBYY FineReader Engine
        UnloadEngine()

def LoadEngine():
    global Engine
    global EngineLoader
    
    DisplayMessage("Initializing Engine...")
    EngineLoader = win32c.Dispatch("FREngine.OutprocLoader.12")
    
    Engine = EngineLoader.InitializeEngine( SamplesConfig.GetCustomerProjectId(), SamplesConfig.GetLicensePath(), SamplesConfig.GetLicensePassword(), "", "", False )
    logging.info('Engine initialized')
    
def ProcessWithEngine(file_path,output_path):
    try:
        ## Setup FREngine
        SetupFREngine()
        ## Process sample image
        ProcessImage(file_path,output_path)
    except Exception as e:
        DisplayMessage( e )

def SetupFREngine():
    global Engine
    
    DisplayMessage( "Loading predefined profile..." )
    Engine.LoadPredefinedProfile( "DocumentConversion_Accuracy" )
    ## Possible profile names are:
    ## "DocumentConversion_Accuracy", "DocumentConversion_Speed",
    ## "DocumentArchiving_Accuracy", "DocumentArchiving_Speed",
    ## "BookArchiving_Accuracy", "BookArchiving_Speed",
    ## "TextExtraction_Accuracy", "TextExtraction_Speed",
    ## "FieldLevelRecognition",
    ## "BarcodeRecognition_Accuracy", "BarcodeRecognition_Speed",
    ## "HighCompressedImageOnlyPdf",
    ## "BusinessCardsProcessing",
    ## "EngineeringDrawingsProcessing",
    ## "Version9Compatibility",
    ## "Default"

def ProcessImage(file_path,output_path):
    global Engine

    imagePath = file_path #os.path.join(SamplesConfig.GetSamplesFolder(),"Input\\co_p001aft183.pdf")
   
    ## Don't recognize PDF file with a textual content, just copy it
    # if( Engine.IsPdfWithTextualContent( imagePath, None ) ):
    #     DisplayMessage( "Copy results..." )
    #     resultPath = os.path.join(SamplesConfig.GetSamplesFolder(),"SampleImages\\Demo_copy.pdf")
    #     copy2( imagePath, resultPath )
    #     return
    
    ## Create document
    document = Engine.CreateFRDocument()

    try:
	## Add image file to document
        DisplayMessage( "Loading image..." )
        document.AddImageFile( imagePath, None, None )

	## Process document
        DisplayMessage( "Process..." )
        document.Process( None )

	## Save results
        DisplayMessage( "Saving results..." )
        FEF_RTF = 0
        FEF_PDF = 4
        FEF_DOCX = 8
        PES_Balanced = 1
        
        export_params = Engine.CreateRTFExportParams()
        export_params.PictureExportParams.Resolution = 300
        export_params.BackgroundColorMode = 1
        export_params.PageSynthesisMode = 2
        export_params.KeepPageBreaks= 1
        export_params.UseDocumentStructure = True

	## Save results to rtf with default parameters
        rtfExportPath = output_path #os.path.join(SamplesConfig.GetSamplesFolder() , "Input\\co_p001aft183.docx")
        document.Export( rtfExportPath,FEF_DOCX , export_params )#FEF_RTF

	# ## Save results to pdf using 'balanced' scenario
    #     pdfParams = Engine.CreatePDFExportParams()
    #     pdfParams.Scenario = PES_Balanced

    #     pdfExportPath = os.path.join(SamplesConfig.GetSamplesFolder() , "SampleImages\\Demo_Python3.pdf")
    #     document.Export( pdfExportPath, FEF_PDF, pdfParams )
    
    except Exception as e:
        DisplayMessage( e )
    finally:
        ## Close document
        document.Close()

def UnloadEngine():
    global Engine
    global EngineLoader
    DisplayMessage( "Deinitializing Engine..." )
    Engine = None
    EngineLoader.ExplicitlyUnload()
    EngineLoader = None

def DisplayMessage( message ,excp_flag = False):
    if excp_flag:
        logging.error(message)
    else:
        logging.info(message)

# try:
#     ## Include config-file SamplesConfig.py
#     # with open('..\\SamplesConfig.py') as f:
#     #     code = compile(f.read(), '..\\SamplesConfig.py', 'exec')
#     #     exec(code)
#     EngineLoader = None
#     Engine = None

#     Run(r"C:\File\NETSCAN\Input\co_a009a116BasisAndPurpose.pdf",r"C:\File\NETSCAN\Input\co_a009a116BasisAndPurpose_pg_frame.docx")
# except Exception as e:
#     DisplayMessage( e )
