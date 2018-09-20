# ExcelImportKit
import excel data(.xlsx file) with epplus by an easy way

## Overview

### Support
.Net Framework 4.5+

### Get Started

1. add *ErrorMessage.cfg.xml* and *ExcelImport.cfg.xml* to *Configs* folder under root directory of project 

2. map importing class fields to *ExcelImport.cfg.xml*

```xml
  <Sample dataStartRow="2" sheetIndex="1" entity="ModelImport.SampleImport, ModelImport" checkEndCol="1">
    <column name="Age" property="Age" col="1" type="System.Int32" regexp="\d+" />
    <column name="Name" property="Name" col="2" type="System.String" unique="true"/>
    <column name="Birthday" property="Birthday" col="3" type="System.Nullable`1[[System.DateTime]]" />
    <column name="Height" property="Height" col="4" type="System.Single" />
    <column name="Money" property="Money" col="5" type="System.Decimal" />
    <column name="Gender" property="Gender" col="6" type="System.String" valuemapping="true" valuetype="System.Boolean">
      <mappings>
        <!--default value-->
        <mapping key="" value="false"/>
        <mapping key="MALE" value="true"/>
        <mapping key="FEMALE" value="false"/>
      </mappings>
    </column>
    <column name="State" property="State" col="7" type="System.String" valuemapping="true" valuetype="System.Int32">
      <mappings>
        <!--default value-->
        <mapping key="" value="0"/>
        <mapping key="STUDENT" value="1"/>
        <mapping key="STAFF" value="2"/>
        <mapping key="SOLDIER" value="3"/>
      </mappings>
    </column>
  </Sample>

```
    node name "Sample" can be any word you like, and all can put multi nodes in cfg file.

3. use it as simple as sample below:

```C#
    IList<ImportError> errors = new List<ImportError>();
    IList<SampleImport> importList;
    var filePath = RootDirectoryHelper.GetFilePath("./Excels/sampleImport.xlsx");
    var cfgNodeName = "Sample";

    using (var fs = new FileStream(filePath, FileMode.Open))
    {
        importList = new ExcelImportService<SampleImport>().GetParsedPositionImport(fs, errors, cfgNodeName);
    } 
```

4. *ErrorMessage.cfg.xml* contains general errors in it. you can custom them of course.

