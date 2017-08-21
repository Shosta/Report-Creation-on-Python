# Report-Creation-on-Python
Generate my Security Report with a Python module.

### 1. Prior to installation
To make this module works properly, you need to install several external modules.

You need :
* [xlxswriter](https://xlsxwriter.readthedocs.io/index.html)
* [xlwt](https://pypi.python.org/pypi/xlwt)
* [python-pptx](https://python-pptx.readthedocs.io/en/latest/#)

The easiest way to do this is to use the *pip* installer.
```python
pip install xlsxwriter
pip install xlwt
pip install python-pptx
```

To develop this program, I use [Visual Studio Code](https://code.visualstudio.com/) with the [Python](https://marketplace.visualstudio.com/items?itemName=donjayamanne.python) extension.

![Visual Studio Code and Python Extension](./ReadmeImages/VisualStudioAndPythonExtension.png)


### 2. Installation

Download the zip file from Github or clone the repository.

### 3. Set up your directories

1. Describe your phases as root directory.
1. Describe the projects and objects you need to follow inside each phase directory.

The directory of project and object must be described as followed : 
 "**_Phase - Project Name - Object Name_**"

 The Phase must be one of the following : 
 * Done
 * In Progress
 * Not Started
 * Not Evaluated

### 4. Launch the application

Then you only need to run the *writereport.py* module in your terminal.

`py ./writereport.py`

### 5. Explanation about creating command alias on Windows, MacOS and Linux

For Windows :
https://technet.microsoft.com/en-us/library/ee176913.aspx

change the Execution policy if you don't have any rights to launch powershell profiles :
https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-5.1&viewFallbackFrom=powershell-Microsoft.PowerShell.Core
