## Common VBA Log Service
### Usage at a glance
The below provides a first impression of the Class-Module's methods and properties
```vb
    Dim Log As New clsLog
    '~~ Preparation
    With Log
        .WithTimeStamp = True               ' defaults to False when ommited
        .AlignmentItems "|C|L.:|L|"         ' explicit items alignment spec
        .MaxItemLengths 6, 15, 30           ' explicit spec of the minimum required column width
        .Headers "| Nr | Item | Comment |"  ' implicitly aligned centered
        .ColsDelimiter = " "                ' would default to | otherwise since headers are specified
    End With
    
    '~~ Any code
    
    Log.Entry "xxxx", "yyyyyy", "zzzzzzzz"
    Log.Entry "xxx", "yyyyyyyyyyyyyyy", "zzzzzzzzzzzzzzzzzzzzz"
    Log.Dsply
```
Displays the following log-file entries:
```
23-05-31-20:37:02 =========================================================
23-05-31-20:37:02    Nr          Item                   Comment            
23-05-31-20:37:02  ------ ------------------ ------------------------------
23-05-31-20:37:02   xxxx  yyyyyy ..........: zzzzzzzz
23-05-31-20:37:02   xxx   yyyyyyyyyyyyyyy .: zzzzzzzzzzzzzzzzzzzzz
```
Notes:
1. The top line delimits a previous series of log entries, not written when the log-file is new
2. The specification of the header would default to a columns delimiter | (vertical bar) - which would be inappropriate in the context of the special left alignment spec for the 2nd column
3. The specific alignment for the 2nd column increased the column width from 15 to 18

### Methods
| Method Name              | Function |
|--------------------------|----------|
|`AlignmentHeaders`    | ParamArray of string expressions, explicitly specifies the [alignment](#explicit-headers-alignment-specification) for each column's header.|
|`AlignmentItems`      | ParamArray of string expressions, explicitly specifies the [alignment](#explicit-items-alignment-specification) for each column's item.|
|`ColsItems`           | Writes items provided as ParamArray arranged in columns to the log file. See property `ColsSpecs` for how to specify corresponding columns alignment and width.|
|`Dsply`               | Displays the log-file by means of the application associated with the file's extension, which defaults to .log|
|`Entry`               | Writes a log entry which is either a single string or a number of items, whereby the latter indicated that the items are to be aligned in columns in accordance with the specified `AlignmentItems`.|
|`Headers`             | ParamArray of string expressions, specifies a header line with column headers. The method may be repeated for multiple column headers. The specified headers may implicitly specify the column headers' [alignment](#implicit-headers-alignment-specification) |
|`MaxItemLengths`      | ParamArray of integer expressions, specifies the [maximum length of items ](#the-maximum-items-length-specification) aligned in columns.|
|`NewFile`             | Creates a new log-file. Mainly used with regression test. |
|`NewLog`              | Explicit indication that the next Entry is the first of a new series of log entries. I.e. with the next Entry a Delimiter line is written (provided its not a new log file), a title is written (provided one has been specified), and a header is written (provided one has been specified). An explicit call of this method is only required, in case neither a Title nor Headers had been specified but a delimiter line should indicate a new series of log entries. `NewLog False` suppresses the delimiter line.|
|`Title`               | ParamArray of strings, each representing a title line (alternatively the method may be called for each line). The alignment of the title lines may implicitly be specified with the first string/method call (see).|
|`WithTimeStamp`       | When not called log entry lines are written without a time-stamp, else with a leading timestamp in the format `yy-mm-dd-hh:mm:ss` |

### Properties
| Name          | Description |
|---------------|-------------|
|`ColsDelimiter`| String expression, read/write, defaults to a single space. <br>Read: returns the current specified `ColsDelimiter`.<br>Write: specifies one. When `|` (vertical bar) is specified, the alignment margin defaults to a single space, when a single space is specified the alignment margin defaults to a vbNullString.|
|`ColsSpecs`    | String expression, specifies each column's alignment (`L`eft, `C`entered, `R`ight, its width and its fill if any. Fill by default are spaces, alternatives are:<br>`"-"` dash<br>`" -"` space dash (starts with a space followed by dashes)<br>`"."` dot<br>`" ."` space dot<br>|
|`FileBaseName` | String expression, write only, specifies the log-file's name, defaults to the  `ActiveWorkbook's` [^1] `BaseName` with an `.log` file extension. |
|`FileExtension`| String expression, read/write, specifies the log-file's file extension, defaults to `log`|
|`FileFullName` | String expression, read/write, write: specifies the full name of the log-file, replaces any defaults (location, base-name, extension; read: returns the joint `FileLocation\FileBaseName.FileExtension`.|
|`FileLocation` | String expression, read/write, specifies the log-files location, defaults to the `ActiveWorkbook.Path`.|
|`KeepLogs`     | Integer expression, write only, defaults to 1 when not provided, specifies the number of logs (lines between delimiter lines) are kept in the log-file.<br><u>. |
|`LogFile`      | File object representing the current active log-file. |

### Installation
1. Download the [clsLog][1] Class-Module and import it into your VB-Project. [^2]

### Columns alignment specifics
The alignment (explicit or implicit) of log items in columns - and optional column headers - is the main focus of the VBA-Log-Service whereby an explicit (`AlignmentHeaders`/`AlignmentItems`) specification takes precedence over an implicit (`Headers`/`Entry`) specification disregarding which method has been called first. I e. for example: When the `Headers` method implicitly specifies the alignment this alignment would be overwritten by a subsequent `AlignmentHeaders` specification.

#### Columns delimiter
When column `Headers` are specified the delimiter defaults to a  `|` (vertical bar), else to a single space.

#### Columns margin (depending on the columns delimiter)

| Columns Delimiter        | Columns Margin              | Comment |
|--------------------------|-----------------------------|---------|
|<nobr>`"\|"` (vertical bar)|<nobr>" " (single space)     | Default when `Headers` are specified. The column content will have at least one a leading and one trailing space. |
| " " (single space)       |<nobr> "" (`vbNullString`)   | Default when no `Headers` are specified. The column content will have no leading and trailing space other than the single space column delimiter. |


#### Columns Width
The column width is the space between two [column delimiters](#columns-delimiter) which may be a `|` (vertical bar) or a single space. The final width of a column considers:
 - the `MaxItemLengths` (when specified for the column)
 - a leading and trailing single space when the `ColsDelimiter` is a vertical bar - which is the default when `Headers` were specified
 - the width of the columns `Headers` (when specified)
 - the length of the first `Entry` item
 - the special left [alignment](#alignment-specification) fill option `.:` which extends the width by 3.
 
 Examples with a `MaxItemLengths` = 10 explicitly specified for a column's items:

| Example | Conditions | Final Cols Width |
| --------|------------|:----------------:|
|<nobr>`| xxxxxxxxxx |`| The cols delimiter is a `|` (vertical bar) - the default either when `Headers` are specified of when the `ColsDelimiter` explicitly specifies it. In both cases the columns margin is a single leading and trailing space.| `MaxItemLengths` + 2 |
|<nobr>` xxxxxxxxxx `| The columns delimiter is a single space, either by default when no `Headers` are specified or when explicitly specified by the `ColsDelimiter`| `MaxItemLengths` |
|<nobr>` xxxxxxxxxx .: `| - The columns delimiter is a ` `(single space), either by default when no `Headers` are specified or when explicitly specified by the `ColsDelimiter`<br>- The `AlignmentIems` specified `"L.:"` for the column indicating fill with `.`(dots) terminated by a `:`(colon) | `MaxItemLengths` + 3 |

#### The maximum items' length specification
The explicit specification of the maximum items' length ensures enough column space in case the the first `Entry`'s items not implicitly specifies them - rarely the case in practice. The specified [`MaxItemLengths`](#methods) becomes the minimum width for the corresponding column. This width may still be expanded by the width of the corresponding column's `Headers` (when specified and greater).

| Example                            | Result |
|------------------------------------|--------|
|<nobr>`.MaxItemLengths "|20|30|25|"`| the minimum width for the column 1, 2, and 3 |
|<nobr>`.MaxItemLengths 20, 30, 25`  | same as above. |
|<nobr>`.MaxItemLengths , ,30`       | the minimum column width is only specified for the 2nd column. For all the other columns the width is determined by the `Headers` (when specified) and the width of the very first `Entry` line's items width.|

### Alignment specification
| Explicitly<br>implicitly | for | Method |
|:--------------:|:----------------------------------------------------------:|-------------------|
| **explicitly** |[headers](#explicit-headers-alignment-specification)        | `AlignmentHeaders`|
| **explicitly** |`Entry` [items](#explicit-items-alignment-specification)    | `AlignmentItems`  |
| **implicitly** |[headers](#implicit-headers-alignment-specification)        | `Headers`         |
| **implicitly** |`Entry` [items](#implicit-alignment-specification-rules)    | `Entry`           |
| **implicitly** |log title lines                                             | `Title`           |

#### Explicit headers alignment specification
```vbs
    .AlignmentHeaders "|C|L|R|"     ' centered, left adjusted and right adjusted
    .AlignmentHeaders "C","L","R"   ' same as above
    .AlignmentHeaders "|x|x | x|"   ' same as above, follows the same rules as the implicit alignment soec 
```
For each header column the alignment is not explicitly specified by means of the corresponding method, the alignment follows the [Implicit column width and alignment specification](#implicit-column-alignment-specification). [^2]

##### Explicit items alignment specification
The below illustrates the variants for how the alignment may explicitly be specified by means of the `AlignmentItems` method.

| Alignment spec | Alternative 1 | Alternative 2 | Alternative 3 |
|----------------|---------------|---------------|---------------|
| <nobr>col 1: centered<br>col2: left adjusted<br>col 3: right adjusted | <nobr>`"|C|L|R|"` | <nobr>`"C","L","R"` | `"|x|x | x|"` |
| <nobr>col 1: centered<br>col 2: left adjusted filled with . (dots)<br>col  3: right adjusted | <nobr>`"|C|L.|R|"`| <nobr>`"C","L.","R"`| <nobr>`"|x|x.| x|"`|
|<nobr>col 1: centered<br>col 2: left adjusted filled with . (dots) terminated by a : (colon)<br>col 3: right adjusted |`"|C|L.:|R|"` |`"C","L.:","R"` |`"|x|x.:| x|"` |


#### Implicit `Headers` alignment specification
Appropriate for the `Headers` specifications, possible but less  likely for `Entry` items. The following example uses the implicit alignment spec for a 3 line header and an explicit alignment spec for the `Entry` items, by intention without a `Title` spec:
```vb
    Dim Log As New clsLog
   
    With Log
        '~~ Preparation
        .WithTimeStamp = True                   ' defaults to False when ommited
        .AlignmentItems "|R|C|L|"               ' explicit items alignment spec
        .MaxItemLengths 6, 15, 30               ' explicit spec of the minimum required column width
        .Headers "| Column|  Column  |Column |" ' only this first header line implicitly specifies the alignment
        .Headers "|    1  |   2      |   3   |" ' any alignment implied is ignored
        .Headers "|(right)|(centered)|(left) |" ' any alignment implied is ignored
    End With
    
    '~~ Any code
    
    Log.Entry "xxxx", "yyyyyy", "zzzzzzzz"
    Log.Dsply

End Sub
```
Writes - and displays:
```
23-06-01-15:06:33 =========================================================
23-06-01-15:06:33 |  Column |    Column     | Column                       
23-06-01-15:06:33 |     1   |       2       | 3                            
23-06-01-15:06:33 | (right) |  (centered)   | (left)                       
23-06-01-15:06:33 +---------+---------------+------------------------------
23-06-01-15:06:33 |    xxxx |    yyyyyy     | zzzzzzzz
```

#### Implicit alignment specification rules
Note that an implicit alignment spec is appropriate for the `Headers` specification rather than with the `Entry` items spec.

| Alignment      | Rule | Examples |
|----------------|------|----------|
| Left adjusted  | 1. The number of leading spaces is less than the number of trailing spaces.<br>2. Fills: A trailing `.` (dot) indicates filled with trailing `.` (dots), a trailing `.:`(dot colon) indicates filled with `.` (dots) terminated with a `:`(colon).<br>Note: Leading spaces are preserved. | `"xxxxx "` left adjustes<br>`"xxxxx."` left adjusted, filled with .<br>`"xxxxx.:"`<br>`" xxxxx    "`|
| Centered       | 1. The number of leading and trailing spaces is equal (may be 0)<br>2. Fills: A leading and trailing `-` indicates filled with `-`.<br>Examples: "xxx", " xxxx ", "-xxxxx-", "- xxxxx -"|
| Right adjusted | The number of trailing spaces is less than the number of leading spaces.<br>Note: Trailing spaces are preserved.| `" xxxx"` |


[^1]: When the `ActiveWorkbook` is used as the default for the log-file's location the log-file is located in the serviced Workbook's parent folder. When the service writing the log is for the Workbook itself `ThisWorkbook` and `ActiveWorkbook` are the same, when the service is provided by another Workbook for the  servicing Workbook will be `ThisWorkbook` and the serviced Workbook will be the `ActiveWorkbook`. In both cases the log-file written into the **serviced Workbook's** parent folder.
 
[^2]: The Workbook (its dedicated parent folder respectively) is dedicated to the Class-Module's development and test and provides a full regression test which compares the result of a series of test with a file containing the expected results.

[1]: https://github.com/warbe-maker/VBA-Log-Service/blob/main/CompMan/source/clsLog.cls
