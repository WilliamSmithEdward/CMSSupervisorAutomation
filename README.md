# CMSSupervisorAutomation
VB.Net code for interfacing with Avaya's CMS Supervisor call center management software. Testing was performed using CMS Supervisor R19.

## Example

The example below establishes a connection to both an MS Access database and a CMS server. After setting up the connections, it triggers a report generation based on selected date ranges and finally, disconnects from the CMS server.

* Initializes a connection to an MS Access database using the given path: "PATH_TO_ACCESS_DB.accdb".
* Initializes a connection to the CMS server with specified server details and user credentials.
* Connects to the CMS server.
* Calls the `RunMyCMSReport` method of the Reporting class to generate a report. The report is based on the date range selected using startDateTimePicker and endDateTimePicker controls.
* Disconnects from the CMS server.

```csharp
Private Sub RunReportButton_Click(sender As Object, e As EventArgs) Handles RunReportButton.Click

    Dim access As New MSAccessConnection("PATH_TO_ACCESS_DB.accdb")

    Dim connection As New CMSSupervisorConnection("CMS_SERVER", 2, "Username", "Password")
    connection.Connect()

    Reporting.RunMyCMSReport(connection,
        access, 
        startDateTimePicker.Value.ToString("M/d/yyyy"),
        endDateTimePicker.Value.ToString("M/d/yyyy")
    )

    connection.Disconnect()

End Sub
```

### Considerations
* Ensure the path to the MS Access database ("PATH_TO_ACCESS_DB.accdb") is valid and accessible.
* The logic in the `RunMyCMSReport` method should be tailored to your CMS report as well as the database table you have implemented.
* The server address, ACD number, and user credentials for the CMS server should be valid to establish a successful connection.
* Date range is selected using startDateTimePicker and endDateTimePicker controls, and it's essential they are correctly set before triggering this method.

## CMSSupervisorConnection Class

This class is used to establish and manage a connection with the CMS server.

### Properties

- **_serverAddress** (String): Address of the server.
- **_acd** (Integer): ACD number.
- **_userName** (String): User name for connection.
- **_password** (String): Password for connection.
- **_cvsApp** (ACSUP.cvsApplication): CVS Application instance.
- **_cvsConn** (ACSCN.cvsConnection): CVS Connection instance.
- **_cvsSrv** (ACSUPSRV.cvsServer): CVS Server instance.

### Constructor

#### `Public Sub New(ByVal serverAddress As String, ByVal acd As Integer, ByVal userName As String, ByVal password As String)`

Constructor for the CMSSupervisorConnection class. Initializes any default values, sets up required configurations, and prepares the instance for communication with the CMS server.

* **serverAddress**: The hostname or IP address of the CMS server.
* **acd**: The Automatic Call Distributor (ACD) number. Represents the ACD system number on the CMS which the connection is intended for.
* **userName**: The username used for authentication.
* **password**: The password used for authentication.

### Methods

#### `Public Sub Connect()`

Establishes the connection to the CMS server using provided credentials.

#### `Public Sub Disconnect()`

Disconnects and releases the resources associated with the CMS server connection.

#### `Public Function ExecuteQuery(reportPath As String, reportParams As Dictionary(Of String, String), timeZone As String) As String`

Executes a report query on the CMS server and returns the result as a string.

- **reportPath**: Path to the desired report.
- **reportParams**: Key-value pairs of parameters for the report.
- **timeZone**: Time zone for the report (pass empty string if the report does not implement time-zone).

## MSAccessConnection Class
This class provides a connection to a Microsoft Access database.

### Properties
* **_dbFilePath** : A string indicating the path to the Microsoft Access database file.

### Methods

#### `Public Function QueryIntoDataTable(sql As String) As DataTable`

Connects to the MS Access database and fetches the data as per the provided SQL query into a DataTable.

* **sql**: A SQL query string to fetch data from the database.
* **Returns**: A DataTable object containing the result of the query.

#### `Public Sub ExecuteSql(sql As String)`

Connects to the MS Access database and executes the provided SQL command, which is typically used for non-query operations like Insert, Update, and Delete.

* **sql**: A SQL command string for database operation.

## Reporting Class

This class provides example methods for generating different types of reports.

### Methods

#### `Public Shared Sub RunSkillSummary(cms As CMSSupervisorConnection, access As MSAccessConnection, startDate As String, endDate As String)`

This method is designed to process and insert skill summary reports into an MS Access database from the data extracted using a CMS server connection.

* **cms**: CMSSupervisorConnection object.
* **access**: MSAccessConnection object.
* **startDate**: Start date for the report.
* **endDate**: End date for the report.

1.) Date Loop:  
* The method starts by setting the sDate (current processing date) to startDate.
* There's a loop that processes data for each date from startDate to endDate.
    
2.) Data Extraction and Insertion:
* For each date in the loop:
    * The method checks if there's already existing data for the current date (sDate) in the Access database. If not:
        * It fetches the skill summary data for the current date from the CMS server using the cms.ExecuteQuery method.
        * This data is then parsed and split into individual lines.
        * The relevant line containing data fields is then split into individual fields.
        * An SQL INSERT statement is constructed using these fields and then executed to insert the data into the DATA_Skill_Summary table in the Access database.

3.) Date Increment:
* After processing for the current date is completed, sDate is incremented by 1 day.
* The loop continues until all dates from startDate to endDate have been processed.

4.) After all dates have been processed, a final SQL statement (Execute []) is executed on the Access database.
