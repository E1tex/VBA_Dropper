# VBA_Dropper
This VBA_Dropper facilitates the automation of downloading a ZIP archive from a server, extracting its contents into a temporary folder, and executing the extracted executable file. Specifically tailored for integration into Microsoft Office macros, it enhances efficiency in security testing scenarios.

It can be utilized for various purposes, including penetration testing and Red Team assessments.
## Requirements:

    - Microsoft Office Package installed.
    - Access privileges to the server hosting the ZIP archive.
    - Enabled "Microsoft Shell Controls And Automation" reference(Macros Project Window -> Tools -> References)
    
## Usage Instructions:

    Configuration:
        Modify the url variable within the DownloadUnzipAndRun subroutine to point to the desired location of the ZIP archive on the server.
        Adjust the server IP address, port, and path to the ZIP file as required.

    Integration:
        Incorporate the provided VBA code into your Microsoft Office macros project.

    Execution:
        The macro will automatically trigger upon opening the document (Document_Open event) or can be executed manually (AutoOpen subroutine).
        Ensure appropriate permissions are granted to download files, extract archives, and execute executables from external sources.

## Disclaimer:
This macro dropper script is provided as-is without any warranty. Use it responsibly and ensure compliance with legal and ethical standards. Users assume all risks associated with its usage. Obtain proper authorization before conducting any security assessments.

## License:
Licensed under the [MIT](https://opensource.org/licenses/MIT) License. Refer to the LICENSE file for detailed information.
