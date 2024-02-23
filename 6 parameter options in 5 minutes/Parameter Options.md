# 6 Popular Paramter Options in 5 Minutes
- **Mandatory**: This parameter option indicates that when calling the function or cmdlet, a specific parameter must be provided.
    ```PowerShell
    [Parameter(Mandatory=$true)]
    ```
- **ValueFromPipeline**: This option indicates that the function can take its input directly from the pipeline.
    ```PowerShell
    [Parameter(ValueFromPipeline=$true)]
    ```
- **ValueFromPipelineByPropertyName**: It allows a function to take the input from properties of the objects passed via the pipeline that share the parameter name.
    ```PowerShell
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    ```
- **Position**: It is used to allow positional parameters. When position is specified, you can supply the arguments to the function in the order of their position.
    ```PowerShell
    [Parameter(Position=0)]
    ```
- **HelpMessage**: It is a string that provides descriptive help for the parameter, displayed when one uses the Get-Help cmdlet.
    ```PowerShell
    [Parameter(HelpMessage="Enter the time of day.")]
    ```
- **ValidateSet**: This option restricts the parameter input to a set of predefined values. If a user attempts to use a value outside of this set, PowerShell will throw an error. It's useful for creating enumerated lists of allowable inputs
    ```PowerShell
    [ValidateSet("Morning", "Afternoon", "Evening")]
    ```
- Combine them all
  ```PowerShell
    [Parameter( Mandatory=$true, 
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true, 
                Position=0,
                HelpMessage="Enter the time of day.")]
    [ValidateSet("Morning", "Afternoon", "Evening")]
    [string]$ParamName
  ```