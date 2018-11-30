# excellent

> ### UTILS TO **TRANSFORM** / **MERGE** / **DIFF** EXCEL FILES

```bash
# Publish package to nuget.org
nuget push ./bin/excellent.1.0.1.nupkg -ApiKey <key> -Source https://api.nuget.org/v3/index.json

# Install from nuget.org
dotnet tool install -g excellent
dotnet tool install -g excellent --version 1.0.x

# Install from local project path
dotnet tool install -g --add-source ./bin excellent

# Uninstall
dotnet tool uninstall -g excellent
```
> NOTE: If the Tool is not accesible post installation, add `%USERPROFILE%\.dotnet\tools` to the PATH env-var.
---

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

<!-- For Transformation -->
<escapeChars>
  <add key="'" value="''" />
</escapeChars>
```
**`excellent.exe.config`**

##### HELP:
> **`general`**
```batch
  -v, --verbose        Prints all messages to standard output.
  --help               Display this help screen.
  --version            Display version information.
```

> **`transform`**
```batch
  -i, --input          Required. Input file to be transformed.
  -f, --format         Transformation format (using Smart-Format).
  -o, --output         Output file.
  -d, --remove-dups    (Default: true) Remove duplicates rows based on PK
  -c, --ignore-case    (Default: true) Ignore casing while performing comparisons
```

> **`merge`**
```batch
  -i, --input          Required. Input files to be merged.
  -l, --keep-left      Retain the values from Left file when a duplicate exists.
  -r, --keep-right     Retain the values from Right file when a duplicate exists.
  -o, --output         Output file.
  -d, --remove-dups    (Default: true) Remove duplicates rows based on PK
  -c, --ignore-case    (Default: true) Ignore casing while performing comparisons
```

> **`diff`**
```batch
  -i, --input          Required. Input files to be diff'd.
  -s, --sqlconn        Sql Server Connection-string).
  -o, --output         Output file.
  -d, --remove-dups    (Default: true) Remove duplicates rows based on PK
  -c, --ignore-case    (Default: true) Ignore casing while performing comparisons
```

---
#### VS EXCEL-DIFF EXTERNAL GUI TOOL ([USING BEYOND-COMPARE](http://www.scootersoftware.com/support.php?zz=kb_vcs#visualstudio-git))

> **.git\config**
```bash
# 'Program Files' OR 'Program Files (x86)' based on the installation
[diff]
    tool = bc4
[difftool "bc4"]
    cmd = \"C:\\Program Files\\Beyond Compare 4\\BComp.exe\" \"$LOCAL\" \"$REMOTE\"
[merge]
    tool = bc4
[mergetool "bc4"]
    cmd = \"C:\\Program Files\\Beyond Compare 4\\BComp.exe\" \"$REMOTE\" \"$LOCAL\" \"$BASE\" \"$MERGED\"
```
> [**`COMPARING MULTI-SHEET EXCEL FILES`**](https://www.scootersoftware.com/support.php?zz=kb_multisheetexcel)
