# excellent

> ### UTILS TO **TRANSFORM** / **MERGE** / **DIFF** EXCEL FILES

##### SAMPLE USAGE:
```batch
excellent -transform -i "TestData\Localizations_1.xlsx" -o "Localizations_1.sql"
excellent -merge -i "TestData\Localizations_1.xlsx" "TestData\Localizations_2.xlsx"
excellent -diff -i "TestData\Localizations_1.xlsx" "TestData\Localizations_2.xlsx"
```
---
##### CONFIG:
```xml
<add key="PrimaryKey" value="{ResourceId}_{ResourceSet}" />
<add key="TransformFormat" value="EXEC [dbo].[usp_InsertLocalizationData] @ResourceId = '{ResourceId}', @English = '{English}', @French = '{French}', @Spanish = '{Spanish}', @ResourceSet = '{ResourceSet}'" />
<add key="IgnoreCase" value="true" />
```
**`excellent.exe.config`**