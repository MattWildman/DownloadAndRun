# Download/update and run any file

A VBScript which downloads any file (or updates a file if it already exists) and then runs it.

*Windows 7+ only.*

## Usage

*Copy or download 'DownloadAndRunFile.vbs'
*You can pass it four command line arguments: [url of file to download] [location to download to] [custom error message] [custom success message]

### Example usage

    .\DownloadAndRunFile.vbs http://example.com Test.html "Error downloading file." "File downloaded successfully!"
	
If you're connected to the internet then this command will save the homepage of [example.com](http://example.com) as *Test.html* in the same directiory as the script was run from, open a dialogue with the success message, then open *Test.html* in your default HTML viewer.
