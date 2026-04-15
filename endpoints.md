
list templates:

endpoint: data/DefinitionGroupTemplateHeaders

Payload: 
    {
      "@odata.etag":"W/\"JzIwODMyNDgyNjEsNjg3MTk0NzY3MzYn\"","TemplateId":"010 - System Setup","Status":"Validated","Description":"System setup","ValidatedDateTime":"2026-04-14T21:43:44Z"
    }
	
	
list template lines:

endpoint: data/DefinitionGroupTemplateLines

Payload: 
    {
      "@odata.etag":"W/\"JzEsNjg3MTk0NzY3MzYn\"","TemplateId":"010 - System Setup","Entity":"Address and contact information purpose","SysModule":"GAB","LevelInExecutionUnit":10,"FailLevelOnError":"No","Sequence":10,"EntityCategory":"Master","ValidationStatus":"No","Tags":"Address setup","FailExecutionUnitOnError":"No","ExecutionUnit":1
    },
	
	

List DMF projects:	
endpoint: data/DataManagementDefinitionGroups

payload: 
    {
      "@odata.etag":"W/\"JzEsNjg3MTk0NzY3MzYn\"","Name":"010 test","ProjectCategory":"Project","OperationType":"Export","GenerateDataPackage":"No","Description":"","TruncateEntityData":"No"
    }
	
lst DMF entities in a DMF project.
endpoint: data/DataManagementDefinitionGroupDetails

Payload:

    {
      "@odata.etag":"W/\"JzUyMjM0OTc2MCw2ODcxOTQ3NjczNic=\"","DefinitionGroupId":"010 test","EntityName":"Address and contact information purpose","RunValidateField":"Yes","LevelInExecutionUnit":10,"InputFilePath":"","RunBusinessValidation":"Yes","EntityCategory":"Master","FailExecutionUnitOnError":"No","SkipStaging":"Yes","QueryForODBC":"","RunBusinessLogic":"Yes","AutoGenerateMapping":"No","QueryData":"SgEvJwAAEQAB5kkBAAAACk3pAwAAhisAAIwrAACIKwAAiysAAAAAhARMAG8AZwBpAHMAdABpAGMAcwBMAG8AYwBhAHQAaQBvAG4AUgBvAGwAZQBFAG4AdABpAHQAeQAAABEAAegDPABMAG8AZwBpAHMAdABpAGMAcwBMAG8AYwBhAHQAaQBvAG4AUgBvAGwAZQBFAG4AdABpAHQAeQBfADEAAADiCTgATABvAGcAaQBzAHQAaQBjAHMATABvAGMAYQB0AGkAbwBuAFIAbwBsAGUARQBuAHQAaQB0AHkAAAAJTegDAADzGQAAAJIEAgARAAEAAP//////////mwT//5oE//8AAAAAAAAB/////wCQBQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==","IsTransformed":"No","ParentEntityName":"","SysModule":"GAB","EntityXMLName":"","Tags":"Address setup","SequenceInLevel":10,"DefaultRefreshType":"FullPush","Disable":"No","ExcelSheetName":"","SampleFilePathOriginal":"","SampleFilePath":"","ExecutionUnit":1,"SourceFormat":"EXCEL","PackageFilePath":"","ValidationStatus":"Yes","FailLevelOnError":"No"
    },