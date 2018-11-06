# excellent

> ### UTILS TO **TRANSFORM** / **MERGE** / **DIFF** EXCEL FILES

##### SAMPLE USAGE:
```batch
excellent -transform -i "TestData\Localizations_1.xlsx" [-o "Localizations_1.sql"] [-f "EXEC [dbo].[usp_InsertLocalizationData] @ResourceId = '{ResourceId}', @English = '{English}'"]
excellent -merge -i "TestData\Localizations_1.xlsx" ["TestData\Localizations_2.xlsx"] [-o Localizations_Merged.xlsx] [-d]
excellent -diff -i "TestData\Localizations_1.xlsx" "TestData\Localizations_2.xlsx" [-c]
excellent -diff -i "TestData\Localizations_1.xlsx" "SELECT * FROM [dbo].[Localizations]" -s "Data Source=server.database.windows.net;Initial Catalog=master;Persist Security Info=True;User ID=userid;Password=pwd"
```
---
##### CONFIG:
```xml
<add key="PrimaryKey" value="{ResourceId}_{ResourceSet}" />
<add key="TransformFormat" value="EXEC [dbo].[usp_InsertLocalizations] @ResourceId = '{ResourceId}', @English = '{English}', @French = '{French}', @Spanish = '{Spanish}', @ResourceSet = '{ResourceSet}'" />
<add key="DiffSqlCommand" value="SELECT * FROM [dbo].[Localizations]" />
```
**`excellent.exe.config`**