# Working with CSVs in PowerShell

PowerShell has some impressive built in utilities for managing and filtering
data sourced from CSV files.  Often, when working with large data sets, it can
be beneficial to reduce the raw source data to the specific fields and
conditions of interest. To do this, we are going to make use of several
PowerShell cmdlets and functions.

## Cmdlets 

> A cmdlet is a lightweight command that is used in the PowerShell environment.
> The PowerShell runtime invokes these **cmdlets** within the context of
> automation scripts that are provided at the command line.  The PowerShell
> runtime also invokes them programmatically through PowerShell APIs.
> [--Microsoft 
> Docs--](https://docs.microsoft.com/en-us/powershell/scripting/developer/)

- **Get-Member** is a cmdlet we will use to describe the output of other
  PowerShell commands. It is aliased as **gm**.
- **Import-Csv** is a cmdlet used to import CSV files into an object format that
  PowerShell can reason with and manipulate.
- **Where-Object** has a lot of utility.  For our purposes, we will use this to
  filter our imported CSVs to return only the information we are interested in.
  This command can be aliased as **where**.
- **Select-Object** is another multipurpose cmdlet that can be applied in a
  variety of ways.  In this instance, we will use **Select-Object** or its
  alias, **select**, to return to us only the columns we are looking for.
- **ConvertTo-Csv** is the cmdlet we will use to take our PowerShell object and
  turn it back into a spreadsheet-friendly CSV file.
- **Out-File** redirects our results from the PowerShell terminal to a file of
  our choosing.
- **Get-Help** (aliased as **help**) can be used with any of the above cmdlets
  (and others) to give us more information on the command we are using.  For
  example, running `Get-Help Out-File` outputs the following:
```
TOPIC
    Windows PowerShell Help System

SHORT DESCRIPTION
    Displays help about Windows PowerShell cmdlets and concepts. 

LONG DESCRIPTION
    Windows PowerShell Help describes Windows PowerShell cmdlets,
    functions, scripts, and modules, and explains concepts, including
    the elements of the Windows PowerShell language.

    Windows PowerShell does not include help files, but you can read the
    help topics online, or use the Update-Help cmdlet to download help files
    to your computer and then use the Get-Help cmdlet to display the help
    topics at the command line.

    You can also use the Update-Help cmdlet to download updated help files
    as they are released so that your local help content is never obsolete. 

    Without help files, Get-Help displays auto-generated help for cmdlets, 
    functions, and scripts.


  ONLINE HELP
    You can find help for Windows PowerShell online in the TechNet Library
    beginning at http://go.microsoft.com/fwlink/?LinkID=108518. 

    To open online help for any cmdlet or function, type:

        Get-Help <cmdlet-name> -Online

  UPDATE-HELP
    To download and install help files on your computer:

       1. Start Windows PowerShell with the "Run as administrator" option.
       2. Type:

          Update-Help

    After the help files are installed, you can use the Get-Help cmdlet to
    display the help topics. You can also use the Update-Help cmdlet to
    download updated help files so that your local help files are always
    up-to-date.

    For more information about the Update-Help cmdlet, type:

       Get-Help Update-Help -Online

    or go to: http://go.microsoft.com/fwlink/?LinkID=210614


  GET-HELP
    The Get-Help cmdlet displays help at the command line from content in
    help files on your computer. Without help files, Get-Help displays basic
    help about cmdlets and functions. You can also use Get-Help to display
    online help for cmdlets and functions.

    To get help for a cmdlet, type:

        Get-Help <cmdlet-name>

    To get online help, type:

        Get-Help <cmdlet-name> -Online    

    The titles of conceptual topics begin with "About_".
    To get help for a concept or language element, type:

        Get-Help About_<topic-name>

    To search for a word or phrase in all help files, type:

        Get-Help <search-term>

    For more information about the Get-Help cmdlet, type:

        Get-Help Get-Help -Online

    or go to: http://go.microsoft.com/fwlink/?LinkID=113316


  EXAMPLES:
      Save-Help              : Download help files from the Internet and saves
                               them on a file share.
      Update-Help            : Downloads and installs help files from the
                               Internet or a file share.
      Get-Help Get-Process   : Displays help about the Get-Process cmdlet.
      Get-Help Get-Process -Online
                             : Opens online help for the Get-Process cmdlet.
      Help Get-Process       : Displays help about Get-Process one page at a time.
      Get-Process -?         : Displays help about the Get-Process cmdlet.
      Get-Help About_Modules : Displays help about Windows PowerShell modules.
      Get-Help remoting      : Searches the help topics for the word "remoting."

  SEE ALSO:
      about_Updatable_Help
      Get-Help
      Save-Help
      Update-Help
```

Now that we have established some of the commands we will be using, let's look
at chaining them together to perform meaningful transformations.

## Chaining commands

PowerShell, like many other scripting languages, uses *pipes* to direct the
output of one command into the input of another.  A pipe is represented as a
vertical line like so: `|`

**Get-Member** accepts, as its input, a PowerShell cmdlet and outputs it
properties and methods.  So, for example, piping **Get-Help** to **Get-Member**
as shown below:

```
Get-Help | Get-Member
```

Outputs the following:

```



   TypeName: System.String

Name             MemberType            Definition                               
----             ----------            ----------                               
Clone            Method                System.Object Clone(), System.Object I...
CompareTo        Method                int CompareTo(System.Object value), in...
Contains         Method                bool Contains(string value)              
CopyTo           Method                void CopyTo(int sourceIndex, char[] de...
EndsWith         Method                bool EndsWith(string value), bool Ends...
Equals           Method                bool Equals(System.Object obj), bool E...
GetEnumerator    Method                System.CharEnumerator GetEnumerator(),...
GetHashCode      Method                int GetHashCode()                        
GetType          Method                type GetType()                           
GetTypeCode      Method                System.TypeCode GetTypeCode(), System....
IndexOf          Method                int IndexOf(char value), int IndexOf(c...
IndexOfAny       Method                int IndexOfAny(char[] anyOf), int Inde...
Insert           Method                string Insert(int startIndex, string v...
IsNormalized     Method                bool IsNormalized(), bool IsNormalized...
LastIndexOf      Method                int LastIndexOf(char value), int LastI...
LastIndexOfAny   Method                int LastIndexOfAny(char[] anyOf), int ...
Normalize        Method                string Normalize(), string Normalize(S...
PadLeft          Method                string PadLeft(int totalWidth), string...
PadRight         Method                string PadRight(int totalWidth), strin...
Remove           Method                string Remove(int startIndex, int coun...
Replace          Method                string Replace(char oldChar, char newC...
Split            Method                string[] Split(Params char[] separator...
StartsWith       Method                bool StartsWith(string value), bool St...
Substring        Method                string Substring(int startIndex), stri...
ToBoolean        Method                bool IConvertible.ToBoolean(System.IFo...
ToByte           Method                byte IConvertible.ToByte(System.IForma...
ToChar           Method                char IConvertible.ToChar(System.IForma...
ToCharArray      Method                char[] ToCharArray(), char[] ToCharArr...
ToDateTime       Method                datetime IConvertible.ToDateTime(Syste...
ToDecimal        Method                decimal IConvertible.ToDecimal(System....
ToDouble         Method                double IConvertible.ToDouble(System.IF...
ToInt16          Method                int16 IConvertible.ToInt16(System.IFor...
ToInt32          Method                int IConvertible.ToInt32(System.IForma...
ToInt64          Method                long IConvertible.ToInt64(System.IForm...
ToLower          Method                string ToLower(), string ToLower(cultu...
ToLowerInvariant Method                string ToLowerInvariant()                
ToSByte          Method                sbyte IConvertible.ToSByte(System.IFor...
ToSingle         Method                float IConvertible.ToSingle(System.IFo...
ToString         Method                string ToString(), string ToString(Sys...
ToType           Method                System.Object IConvertible.ToType(type...
ToUInt16         Method                uint16 IConvertible.ToUInt16(System.IF...
ToUInt32         Method                uint32 IConvertible.ToUInt32(System.IF...
ToUInt64         Method                uint64 IConvertible.ToUInt64(System.IF...
ToUpper          Method                string ToUpper(), string ToUpper(cultu...
ToUpperInvariant Method                string ToUpperInvariant()                
Trim             Method                string Trim(Params char[] trimChars), ...
TrimEnd          Method                string TrimEnd(Params char[] trimChars)  
TrimStart        Method                string TrimStart(Params char[] trimChars)
Category         NoteProperty          string Category=HelpFile                 
Component        NoteProperty          string Component=                        
Functionality    NoteProperty          string Functionality=                    
Name             NoteProperty          string Name=default                      
Role             NoteProperty          string Role=                             
Synopsis         NoteProperty          string Synopsis=SHORT DESCRIPTION        
Chars            ParameterizedProperty char Chars(int index) {get;}             
Length           Property              int Length {get;}                        

```

Notice the *pipe* `|` between `Get-Help` and `Get-Member`.  This tells
`Get-Help` to *pipe* its output to `Get-Member`.  Without the pipe, we would be
calling `Get-Help Get-Member` which would output the following:

```

NAME
    Get-Member
    
SYNOPSIS
    Gets the properties and methods of objects.
    
    
SYNTAX
    Get-Member [[-Name] <System.String[]>] [-Force] [-InputObject 
    <System.Management.Automation.PSObject>] [-MemberType {AliasProperty | 
    CodeProperty | Property | NoteProperty | ScriptProperty | Properties | 
    PropertySet | Method | CodeMethod | ScriptMethod | Methods | 
    ParameterizedProperty | MemberSet | Event | Dynamic | All}] [-Static] 
    [-View {Extended | Adapted | Base | All}] [<CommonParameters>]
    
    
DESCRIPTION
    The `Get-Member` cmdlet gets the members, the properties and methods, of 
    objects.
    
    To specify the object, use the InputObject parameter or pipe an object to 
    `Get-Member`. To get information about static members, the members of the 
    class, not of the instance, use the Static parameter. To get only certain 
    types of members, such as NoteProperties , use the MemberType parameter.
    

RELATED LINKS
    Online Version: https://docs.microsoft.com/powershell/module/microsoft.power
    shell.utility/get-member?view=powershell-5.1&WT.mc_id=ps-gethelp
    Add-Member 

REMARKS
    To see the examples, type: "get-help Get-Member -examples".
    For more information, type: "get-help Get-Member -detailed".
    For technical information, type: "get-help Get-Member -full".
    For online help, type: "get-help Get-Member -online"

```

The utility of both piping commands, and the cmdlet **Get-Member** will become
apparent as we begin building the commands we need to parse and filter our CSVs
in the next section.

## Building our script

Time to start building our script and transforming our CSVs.

First: An example CSV

**example.csv**
```
SKU,Name,Part Number,Qty,Alt1,Alt2,Alt3,Option1,Option2
1001,Red Widget,101-100,3,,,,Color,
1002,Blue Widget,101-100,10,,,,Color,
1003,Yellow Widget,101-100,10,,,,Color,
1004,Tough Cookies,201-100,10,,,,Flavor,
1005,Crumbled Cookies,201-100,10,,,,Flavor,
1006,Ian's Hats,301-100,10,,,,,
```

### Import-Csv

Our first step is going to be getting our CSV *into* PowerShell in a format that
it can read and understand.  For this we use the command **Import-Csv**.  First,
let's check out the help section for this command:
```
help Import-Csv
```
*Remember that **help** is an *alias*, or another name, for **Get-Help**.*
```
NAME
    Import-Csv
    
SYNOPSIS
    Creates table-like custom objects from the items in a comma-separated value 
    (CSV) file.
    
    
SYNTAX
    Import-Csv [[-Path] <System.String[]>] [[-Delimiter] <System.Char>] 
    [-Encoding {ASCII | BigEndianUnicode | Default | OEM | Unicode | UTF7 | 
    UTF8 | UTF32}] [-Header <System.String[]>] [-LiteralPath <System.String[]>] 
    [<CommonParameters>]
    
    Import-Csv [[-Path] <System.String[]>] [-Encoding {ASCII | BigEndianUnicode 
    | Default | OEM | Unicode | UTF7 | UTF8 | UTF32}] [-Header 
    <System.String[]>] [-LiteralPath <System.String[]>] -UseCulture 
    [<CommonParameters>]
    
    
DESCRIPTION
    The `Import-Csv` cmdlet creates table-like custom objects from the items in 
    CSV files. Each column in the CSV file becomes a property of the custom 
    object and the items in rows become the property values. `Import-Csv` works 
    on any CSV file, including files that are generated by the `Export-Csv` 
    cmdlet.
    
    You can use the parameters of the `Import-Csv` cmdlet to specify the column 
    header row and the item delimiter, or direct `Import-Csv` to use the list 
    separator for the current culture as the item delimiter.
    
    You can also use the `ConvertTo-Csv` and `ConvertFrom-Csv` cmdlets to 
    convert objects to CSV strings (and back). These cmdlets are the same as 
    the `Export-CSV` and `Import-Csv` cmdlets, except that they do not deal 
    with files.
    
    If a header row entry in a CSV file contains an empty or null value, 
    PowerShell inserts a default header row name and displays a warning message.
    
    `Import-Csv` uses the byte-order-mark (BOM) to detect the encoding format 
    of the file. If the file has no BOM, it assumes the encoding is UTF8.
    

RELATED LINKS
    Online Version: https://docs.microsoft.com/powershell/module/microsoft.power
    shell.utility/import-csv?view=powershell-5.1&WT.mc_id=ps-gethelp
    ConvertFrom-Csv 
    ConvertTo-Csv 
    Export-Csv 
    Get-Culture 

REMARKS
    To see the examples, type: "get-help Import-Csv -examples".
    For more information, type: "get-help Import-Csv -detailed".
    For technical information, type: "get-help Import-Csv -full".
    For online help, type: "get-help Import-Csv -online"
```

For our purposes, we are only concerned with a portion of the syntax for this
command: `Import-Csv [[-Path] <System.String[]>]`

The words preceded with a `-` are what are called *flags*.  This either modify
our command or indicate that the next word to follow will modify our command.
In the case of `Import-Csv`, the flag `-Path` is a *positional parameter*
meaning that, even if we omit the `-Path`, PowerShell will assume that the first
expression to follow the command `Import-Csv` will be the path.  Let's try it
out.

```
Import-Csv .\examples\csv\example.csv
```
> We are using the *relative path* here (the path in relation to where we
> currently are in the directory), but we could just as easily use the *literal
> path* which might look something like:
> `Import-Csv
> C:\Users\ian.goodnight\Downloads\Guides\examples\csv\example.csv`
> Generally speaking, PowerShell and other terminals make entering these longs
> paths simpler through something called *tab completion*, meaning that after
> entering `C:\Us` pressing <tab> autocompletes the path so that it now displays
> `C:\Users\` and, in my case, pressing <tab> after `C:\Users\i` autocompletes
> to `C:\Users\ian.goodnight\` and so on.

**Output**

```


SKU         : 1001
Name        : Red Widget
Part Number : 101-100
Qty         : 3
Alt1        : 
Alt2        : 
Alt3        : 
Option1     : Color
Option2     : 

SKU         : 1002
Name        : Blue Widget
Part Number : 101-100
Qty         : 10
Alt1        : 
Alt2        : 
Alt3        : 
Option1     : Color
Option2     : 

SKU         : 1003
Name        : Yellow Widget
Part Number : 101-100
Qty         : 10
Alt1        : 
Alt2        : 
Alt3        : 
Option1     : Color
Option2     : 

SKU         : 1004
Name        : Tough Cookies
Part Number : 201-100
Qty         : 10
Alt1        : 
Alt2        : 
Alt3        : 
Option1     : Flavor
Option2     : 

SKU         : 1005
Name        : Crumbled Cookies
Part Number : 201-100
Qty         : 10
Alt1        : 
Alt2        : 
Alt3        : 
Option1     : Flavor
Option2     : 

SKU         : 1006
Name        : Ian's Hats
Part Number : 301-100
Qty         : 10
Alt1        : 
Alt2        : 
Alt3        : 
Option1     : 
Option2     : 

```

We can see that PowerShell read our CSV without complaint, but it doesn't look
the way we expect.

PowerShell deals with *objects* and so, the output of out call to `Import-Csv`
is formatted in a generic list view as a PowerShell default.

### Get-Member

While **Get-Member** will not be part of our final script, I find that using it
as we build scripts helps to visualize the object that we are dealing with in
the moment.  Let's try our previous command, but this time, we'll *pipe* the out
put to `Get-Member`.

> When repeating commands in PowerShell or other terminals, it can be convenient
> to use the ↑ arrow to cycle through our command history.  In this case,
> pressing ↑ brings up our previous command and allows us to append further
> commands to it.

```
Import-Csv .\examples\csv\example.csv | Get-Member
```

**Output**

```


   TypeName: System.Management.Automation.PSCustomObject

Name        MemberType   Definition                    
----        ----------   ----------                    
Equals      Method       bool Equals(System.Object obj)
GetHashCode Method       int GetHashCode()             
GetType     Method       type GetType()                
ToString    Method       string ToString()             
Alt1        NoteProperty string Alt1=                  
Alt2        NoteProperty string Alt2=                  
Alt3        NoteProperty string Alt3=                  
Name        NoteProperty string Name=Red Widget        
Option1     NoteProperty string Option1=Color          
Option2     NoteProperty string Option2=               
Part Number NoteProperty string Part Number=101-100    
Qty         NoteProperty string Qty=3                  
SKU         NoteProperty string SKU=1001               

```

As before, we see several *methods* defined (which will not utilize for the
purposes of this guide), but, interestingly, we also see several instances of
the MemberType **NoteProperty**.  These *NoteProperties* correspond with our CSV
headers and will be useful later on.

### Where-Object

OK, we we have *imported* a CSV with `Import-Csv` so that PowerShell can reason
with and display the information contained therein, but we haven't done much to
transform the output yet.  Let's look at filtering our output with
**Where-Object**.

First, let's check out the help pages for **Where-Object**.

```
help Where-Object
```
> Remember `help` is an *alias* for `Get-Help`.

**Output**

```

NAME
    Where-Object
    
SYNOPSIS
    Selects objects from a collection based on their property values.
    
    
SYNTAX
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -CContains [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -CEQ [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -CGE [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -CGT [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -CIn [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -CLE [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -CLike [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -CLT [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -CMatch [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -CNE [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -CNotContains [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -CNotIn [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -CNotLike [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -CNotMatch [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -Contains [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] [-EQ] [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-FilterScript] <System.Management.Automation.ScriptBlock> [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -GE [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -GT [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] -In [-InputObject 
    <System.Management.Automation.PSObject>] [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] [-InputObject 
    <System.Management.Automation.PSObject>] -Is [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] [-InputObject 
    <System.Management.Automation.PSObject>] -IsNot [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] [-InputObject 
    <System.Management.Automation.PSObject>] -LE [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] [-InputObject 
    <System.Management.Automation.PSObject>] -Like [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] [-InputObject 
    <System.Management.Automation.PSObject>] -LT [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] [-InputObject 
    <System.Management.Automation.PSObject>] -Match [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] [-InputObject 
    <System.Management.Automation.PSObject>] -NE [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] [-InputObject 
    <System.Management.Automation.PSObject>] -NotContains [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] [-InputObject 
    <System.Management.Automation.PSObject>] -NotIn [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] [-InputObject 
    <System.Management.Automation.PSObject>] -NotLike [<CommonParameters>]
    
    Where-Object [-Property] <System.String> [[-Value] <System.Object>] [-InputObject 
    <System.Management.Automation.PSObject>] -NotMatch [<CommonParameters>]
    
    
DESCRIPTION
    The `Where-Object` cmdlet selects objects that have particular property values from the collection of objects 
    that are passed to it. For example, you can use the `Where-Object` cmdlet to select files that were created 
    after a certain date, events with a particular ID, or computers that use a particular version of Windows.
    
    Starting in Windows PowerShell 3.0, there are two different ways to construct a `Where-Object` command.
    
    - Script block . You can use a script block to specify the property name, a comparison operator,   and a 
    property value. `Where-Object` returns all objects for which the script block statement is   true.
    
    For example, the following command gets processes in the Normal priority class, that is, processes   where the 
    value of the PriorityClass property equals Normal.
    
    `Get-Process | Where-Object {$_.PriorityClass -eq "Normal"}`
    
    All PowerShell comparison operators are valid in the script block format. For more information   about 
    comparison operators, see about_Comparison_Operators (./About/about_Comparison_Operators.md).
    
    - Comparison statement . You can also write a comparison statement, which is much more like   natural 
    language. Comparison statements were introduced in Windows PowerShell 3.0.
    
    For example, the following commands also get processes that have a priority class of Normal. These   commands 
    are equivalent and can be used interchangeably.
    
    `Get-Process | Where-Object -Property PriorityClass -eq -Value "Normal"`
    
    `Get-Process | Where-Object PriorityClass -eq "Normal"`
    
    Starting in Windows PowerShell 3.0, `Where-Object` adds comparison operators as parameters in a   
    `Where-Object` command. Unless specified, all operators are case-insensitive. Prior to Windows   PowerShell 
    3.0, the comparison operators in the PowerShell language could be used only in script   blocks.
    
    When you provide a single Property to `Where-Object`, the value of the property is treated as a boolean 
    expression. When the value of Length is not zero, the expression evaluates to True . For example: `('hi', '', 
    'there') | Where-Object Length`
    
    The previous example is functionally equivalent to:
    
    - `('hi', '', 'there') | Where-Object Length -GT 0`
    
    - `('hi', '', 'there') | Where-Object {$_.Length -gt 0}`
    

RELATED LINKS
    Online Version: https://docs.microsoft.com/powershell/module/microsoft.powershell.core/where-object?view=powers
    hell-5.1&WT.mc_id=ps-gethelp
    Compare-Object 
    ForEach-Object 
    Group-Object 
    Measure-Object 
    New-Object 
    Select-Object 
    Sort-Object 
    Tee-Object 

REMARKS
    To see the examples, type: "get-help Where-Object -examples".
    For more information, type: "get-help Where-Object -detailed".
    For technical information, type: "get-help Where-Object -full".
    For online help, type: "get-help Where-Object -online"


```

We are going to use a *script block* to filter out unwanted input coming from
out `Import-Csv` command.  A call to `Where-Object` with a script block might
look like this:

```
Where-Object { $_.Name -eq "Zaphod Beeblebox" }
```

Note that the *script block* is a *positional parameter* meaning that the script
expects the expression immediately following the call to `Where-Object` to be a
script block.

Script blocks begin and end with a curly brace.

`{  }`

With PowerShell, the variable `$_` refers to *this* record (or *this* row in the
case of our CSV).

`-eq` is a comparison operator (think of it as 'equals').  Try running the
command `help comparison` to get a full list of comparison operators supported
with PowerShell.

Our call to `Where-Object` above won't function correctly, as it is not
receiving and object as input.  Even if it were, our CSV doesn't seem to have a
name propery.  Let's use `Get-Member` on the output of our call to `Import-Csv`
one more time to remind of us of the properties we have available.

```
Import-Csv .\examples\csv\example.csv | gm
```
> **gm** is an *alias* for **Get-Member**

**Output**

```



   TypeName: System.Management.Automation.PSCustomObject

Name        MemberType   Definition                    
----        ----------   ----------                    
Equals      Method       bool Equals(System.Object obj)
GetHashCode Method       int GetHashCode()             
GetType     Method       type GetType()                
ToString    Method       string ToString()             
Alt1        NoteProperty string Alt1=                  
Alt2        NoteProperty string Alt2=                  
Alt3        NoteProperty string Alt3=                  
Name        NoteProperty string Name=Red Widget        
Option1     NoteProperty string Option1=Color          
Option2     NoteProperty string Option2=               
Part Number NoteProperty string Part Number=101-100    
Qty         NoteProperty string Qty=3                  
SKU         NoteProperty string SKU=1001               

```

And, as a reference, let's take a look at our CSV one more time.

```
SKU,Name,Part Number,Qty,Alt1,Alt2,Alt3,Option1,Option2
1001,Red Widget,101-100,3,,,,Color,
1002,Blue Widget,101-100,10,,,,Color,
1003,Yellow Widget,101-100,10,,,,Color,
1004,Tough Cookies,201-100,10,,,,Flavor,
1005,Crumbled Cookies,201-100,10,,,,Flavor,
1006,Ian's Hats,301-100,10,,,,,
```

We can see that all of our "Cookies" share a common "Part Number" (201-100).
How might we write a script that returns to us only the information we are
interested in, in this case, our cookie inventory items?

**The Command**

```
Import-Csv  .\examples\csv\example.csv | Where-Object { $_."Part Number" -eq
"201-100"}
```
> Note that following our *this* character, `$_` we access one of our
> *NoteProperties* (headers) with a `.`
> The property is wrapped in quotes to deal with the white space between "Part"
> and "Number" as this would cause an error.  Using the comparison operator
> `-eq`, we are searching for "Part Numbers" that *equal* "201-100"

**Output**

```


SKU         : 1004
Name        : Tough Cookies
Part Number : 201-100
Qty         : 10
Alt1        : 
Alt2        : 
Alt3        : 
Option1     : Flavor
Option2     : 

SKU         : 1005
Name        : Crumbled Cookies
Part Number : 201-100
Qty         : 10
Alt1        : 
Alt2        : 
Alt3        : 
Option1     : Flavor
Option2     : 

```

As we can see from the output, we are now only displaying information relating
to the products we are interested in.  In this case, cookies.  Let's run a
similar command, but this time, let's look at cookies AND widgets.

```
Import-Csv .\examples\csv\example.csv | Where-Object { $_."Part Number" -in
"101-100","201-100" }
```
> The comparison operator `-in` accepts a comma separated list of comparisons
> and returns any items that matches at least one.

**Output**

```


SKU         : 1001
Name        : Red Widget
Part Number : 101-100
Qty         : 3
Alt1        : 
Alt2        : 
Alt3        : 
Option1     : Color
Option2     : 

SKU         : 1002
Name        : Blue Widget
Part Number : 101-100
Qty         : 10
Alt1        : 
Alt2        : 
Alt3        : 
Option1     : Color
Option2     : 

SKU         : 1003
Name        : Yellow Widget
Part Number : 101-100
Qty         : 10
Alt1        : 
Alt2        : 
Alt3        : 
Option1     : Color
Option2     : 

SKU         : 1004
Name        : Tough Cookies
Part Number : 201-100
Qty         : 10
Alt1        : 
Alt2        : 
Alt3        : 
Option1     : Flavor
Option2     : 

SKU         : 1005
Name        : Crumbled Cookies
Part Number : 201-100
Qty         : 10
Alt1        : 
Alt2        : 
Alt3        : 
Option1     : Flavor
Option2     : 

```

No we are seeing the results of rows where the part number is "101-100" OR
"201-100".  But, what about those empty fields, "Alt1", "Alt2", "Alt3", and
"Option2"?

### Select-Object

Often, when dealing with raw CSV output from services we may subsribe to, the
CSVs generated include extra, unneccessary or unwanted information.  Let's add
on to our script to filter out this extra information.

Let's got back to just looking at cookies, for simplicity's sake.

```
Import-Csv .\examples\csv\example.csv | where { $_."Part Number" -eq "201-100" }
```
> **where** is an *alias* for **Where-Object**

For reference, here is our CSV once again.

```
SKU,Name,Part Number,Qty,Alt1,Alt2,Alt3,Option1,Option2
1001,Red Widget,101-100,3,,,,Color,
1002,Blue Widget,101-100,10,,,,Color,
1003,Yellow Widget,101-100,10,,,,Color,
1004,Tough Cookies,201-100,10,,,,Flavor,
1005,Crumbled Cookies,201-100,10,,,,Flavor,
1006,Ian's Hats,301-100,10,,,,,
```

We've seen how to filter out rows to pare our raw CSV down to the pertinent
information, but what about all those extra blank columns?  In our example,
"Alt1", "Alt2", "Alt3", and "Option2" are fields that are included in our
report, but they are all blank.  Maybe they are a new feature that isn't yet
populated, maybe they are data fields that we don't track in our organization.
We can use **Select-Object** to filter them out giving us just the information
we need.

First, the documentation:
```

NAME
    Select-Object
    
SYNOPSIS
    Selects objects or object properties.
    
    
SYNTAX
    Select-Object [[-Property] <System.Object[]>] [-ExcludeProperty <System.String[]>] [-ExpandProperty 
    <System.String>] [-First <System.Int32>] [-InputObject <System.Management.Automation.PSObject>] [-Last 
    <System.Int32>] [-Skip <System.Int32>] [-Unique] [-Wait] [<CommonParameters>]
    
    Select-Object [[-Property] <System.Object[]>] [-ExcludeProperty <System.String[]>] [-ExpandProperty 
    <System.String>] [-InputObject <System.Management.Automation.PSObject>] [-SkipLast <System.Int32>] [-Unique] 
    [<CommonParameters>]
    
    Select-Object [-Index <System.Int32[]>] [-InputObject <System.Management.Automation.PSObject>] [-Unique] 
    [-Wait] [<CommonParameters>]
    
    
DESCRIPTION
    The `Select-Object` cmdlet selects specified properties of an object or set of objects. It can also select 
    unique objects, a specified number of objects, or objects in a specified position in an array.
    
    To select objects from a collection, use the First , Last , Unique , Skip , and Index parameters. To select 
    object properties, use the Property parameter. When you select properties, `Select-Object` returns new objects 
    that have only the specified properties.
    
    Beginning in Windows PowerShell 3.0, `Select-Object` includes an optimization feature that prevents commands 
    from creating and processing objects that are not used.
    
    When you include a `Select-Object` command with the First or Index parameters in a command pipeline, 
    PowerShell stops the command that generates the objects as soon as the selected number of objects is 
    generated, even when the command that generates the objects appears before the `Select-Object` command in the 
    pipeline. To turn off this optimizing behavior, use the Wait parameter.
    

RELATED LINKS
    Online Version: https://docs.microsoft.com/powershell/module/microsoft.powershell.utility/select-object?view=po
    wershell-5.1&WT.mc_id=ps-gethelp
    Group-Object 
    Sort-Object 
    Where-Object 

REMARKS
    To see the examples, type: "get-help Select-Object -examples".
    For more information, type: "get-help Select-Object -detailed".
    For technical information, type: "get-help Select-Object -full".
    For online help, type: "get-help Select-Object -online"

```

**Select-Object** has many useful flags including `-First` which accepts an
integer and acts as a limit to output `-Last` which performs the inverser
operation, and `-Unique` which behaves as you might expect, but the flag we are
interested in is `-Property`.

`-Property` accepts a comma separated list of the properties we are interested
in seeing in our output.  Let's check out what that will look like.

Hitting the ↑ arrow from our PowerShell will return our last command from
history.  In my case, the last command we entered was:

```
Import-Csv .\examples\csv\example.csv | Where-Object { $_."Part Number" -eq
"201-100" }
```

Now we need to pipe that command to the `Select-Object` command so we can select
the properties we are interested in.  What were the properties of this CSV
again?  We can pipe our last command to `Get-Member` (`gm`) to remind ourselves.

```
Import-Csv .\examples\csv\example.csv | Where-Object { $_."Part Number" -eq
"201-100" } | Get-Member
```

**Output:**

```


   TypeName: System.Management.Automation.PSCustomObject

Name        MemberType   Definition                    
----        ----------   ----------                    
Equals      Method       bool Equals(System.Object obj)
GetHashCode Method       int GetHashCode()             
GetType     Method       type GetType()                
ToString    Method       string ToString()             
Alt1        NoteProperty string Alt1=                  
Alt2        NoteProperty string Alt2=                  
Alt3        NoteProperty string Alt3=                  
Name        NoteProperty string Name=Tough Cookies     
Option1     NoteProperty string Option1=Flavor         
Option2     NoteProperty string Option2=               
Part Number NoteProperty string Part Number=201-100    
Qty         NoteProperty string Qty=10                 
SKU         NoteProperty string SKU=1004               

```
> You mah have noticed our headers are not in the same order they appeared in
> our original CSV.  PowerShell alphabetizes our NoteProperties before
> outputting them.

Again, we are only interested in the `NotePropery` MemberTypes for the moment as
these will correspond with our CSV headers (thanks to the output of
`Import-Csv`).

With this reminder still visible in my terminal, pressing ↑ twice (or pressing
it once and removing the last pipe to `Get-Member`) gets us back to our script
so far:

```
Import-Csv .\examples\csv\example.csv | Where-Object { $_."Part Number" -eq
"201-100" }
```

Let's go ahead and pipe the output of this script to `Select-Object` to filter
our output down to ONLY the populated columns (SKU, Name, Part Number, Qty, and
Option1).

```
Import-Csv .\examples\csv\example.csv | Where-Object { $_."Part Number" -eq
"201-100" } | Select-Object -Property SKU,Name,"Part Number",Qty,Option1
```
> Notice we are again wrapping 'Part Number' in quotes so as not to cause
> problems with interpreting white space.  It is also worth noting that our list
> of properties can be in any order and our output will match.

**Output:**
```


SKU         : 1004
Name        : Tough Cookies
Part Number : 201-100
Qty         : 10
Option1     : Flavor

SKU         : 1005
Name        : Crumbled Cookies
Part Number : 201-100
Qty         : 10
Option1     : Flavor

```

We have now removed our extraneous columns.  We can even use the same operation
to change the order of our columns simply by changing the order we list them
following the `-Property` flag.

### ConvertTo-Csv

Often when manipulating CSVs, we do so with the intention of uploading them as
CSVs again into other systems.  What we have done so far has done a lot to clean
up our data, but it has left it in state that isn't easy to import into other
systems.  Let's look at **ConvertTo-Csv** to solve that issue for us.

But first:

```
Get-Help CovertTo-Csv
```

**Output:**

```


NAME
    ConvertTo-Csv
    
SYNOPSIS
    Converts .NET objects into a series of comma-separated value (CSV) strings.
    
    
SYNTAX
    ConvertTo-Csv [-InputObject] <System.Management.Automation.PSObject> [[-Delimiter] <System.Char>] 
    [-NoTypeInformation] [<CommonParameters>]
    
    ConvertTo-Csv [-InputObject] <System.Management.Automation.PSObject> [-NoTypeInformation] [-UseCulture] 
    [<CommonParameters>]
    
    
DESCRIPTION
    The `ConvertTo-CSV` cmdlet returns a series of comma-separated value (CSV) strings that represent the objects that 
    you submit. You can then use the `ConvertFrom-Csv` cmdlet to recreate objects from the CSV strings. The objects 
    converted from CSV are string values of the original objects that contain property values and no methods.
    
    You can use the `Export-Csv` cmdlet to convert objects to CSV strings. `Export-CSV` is similar to `ConvertTo-CSV`, 
    except that it saves the CSV strings to a file.
    
    The `ConvertTo-CSV` cmdlet has parameters to specify a delimiter other than a comma or use the current culture as 
    the delimiter.
    

RELATED LINKS
    Online Version: https://docs.microsoft.com/powershell/module/microsoft.powershell.utility/convertto-csv?view=powers
    hell-5.1&WT.mc_id=ps-gethelp
    ConvertFrom-Csv 
    Export-Csv 
    Import-Csv 

REMARKS
    To see the examples, type: "get-help ConvertTo-Csv -examples".
    For more information, type: "get-help ConvertTo-Csv -detailed".
    For technical information, type: "get-help ConvertTo-Csv -full".
    For online help, type: "get-help ConvertTo-Csv -online"

```

For our purposes, this command will take no arguments or flags, as the piped
output of `Select-Object` will act as the `-InputObject`.

**The Command:**

```
Import-Csv .\examples\csv\example.csv | Where-Object { $_."Part Number" -eq
"201-100" } | Select-Object -Property SKU,Name,"Part Number",Qty,Option1 |
ConvertTo-Csv
```

**Output:**

```
#TYPE Selected.System.Management.Automation.PSCustomObject
"SKU","Name","Part Number","Qty","Option1"
"1004","Tough Cookies","201-100","10","Flavor"
"1005","Crumbled Cookies","201-100","10","Flavor"
```

Awesome!  That is a much cleaner looking CSV!  Of course, it doesn't do us any
good if we can't get it out of our PowerShell terminal...

### Out-File

The last piece of our puzzle is the **Out-File** command which redirects our
piped output from the terminal to a file of our choosing.

```
Get-Help Out-File
```

**Output:**

```

NAME
    Out-File
    
SYNOPSIS
    Sends output to a file.
    
    
SYNTAX
    Out-File [-FilePath] <System.String> [[-Encoding] {ASCII | BigEndianUnicode 
    | Default | OEM | String | Unicode | Unknown | UTF7 | UTF8 | UTF32}] 
    [-Append] [-Force] [-InputObject <System.Management.Automation.PSObject>] 
    [-NoClobber] [-NoNewline] [-Width <System.Int32>] [-Confirm] [-WhatIf] 
    [<CommonParameters>]
    
    Out-File [[-Encoding] {ASCII | BigEndianUnicode | Default | OEM | String | 
    Unicode | Unknown | UTF7 | UTF8 | UTF32}] [-Append] [-Force] [-InputObject 
    <System.Management.Automation.PSObject>] -LiteralPath <System.String> 
    [-NoClobber] [-NoNewline] [-Width <System.Int32>] [-Confirm] [-WhatIf] 
    [<CommonParameters>]
    
    
DESCRIPTION
    The `Out-File` cmdlet sends output to a file. When you need to specify 
    parameters for the output use `Out-File` rather than the redirection 
    operator (`>`).
    

RELATED LINKS
    Online Version: https://docs.microsoft.com/powershell/module/microsoft.power
    shell.utility/out-file?view=powershell-5.1&WT.mc_id=ps-gethelp
    about_Providers 
    about_Quoting_Rules 
    Out-Default 
    Out-GridView 
    Out-Host 
    Out-Null 
    Out-Printer 
    Out-String 
    Tee-Object 

REMARKS
    To see the examples, type: "get-help Out-File -examples".
    For more information, type: "get-help Out-File -detailed".
    For technical information, type: "get-help Out-File -full".
    For online help, type: "get-help Out-File -online"

```

Just as our command begin with `Import-Csv` and a positional path parameter that
pointed at our target, it will end with `Out-File` followed by the path where
our finished file will end.  This is also your chance to name your file!

**The Command:**

```
Import-Csv .\examples\csv\example.csv | Where-Object { $_."Part Number" -eq
"201-100" } | Select-Object -Property SKU,Name,"Part Number",Qty,Option1 |
ConvertTo-Csv | Out-File .\examples\csv\finished.csv
```

This will save the file as "finished.csv" in my CSV directory.

## Going Further

With these tools, we can do a lot to transform, filter, and reduce raw reports
as generated from various applications.  Explore the help documentation through
`Get-Help` further for further commands, techniques, and examples.
