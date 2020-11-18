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
(Microsoft Docs)[https://docs.microsoft.com/en-us/powershell/scripting/developer/]

- **Get-Member** is a cmdlet we will use to describe the output of other
  PowerShell commands. It is aliased as **gm**.
- **Import-Csv** is a cmdlet used to import CSV files into an object format that
  PowerShell can reason with and manipulate.
- **Where-Object** has a lot of utility.  For our purposes, we will use this to
  filter our imported CSVs to return only the information we are interested in.
  This command can be aliased as **where**.
- **Select-Object** is another multipurpose cmdlet that can be applied in a
  variety of ways.  In this instance, we will use **Select-Object** or it's
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


