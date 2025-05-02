Move the whole "Python" folder that contains this ReadMe file to %ProgramData%\ABBYY\SDK\12\FineReader Engine\Samples.

Modify the SamplesConfig.py file: specify the Customer Project ID of your license and, if necessary, the path to the cloud license file and license password.

The code samples require the pywin32 package for correct operation. You can install the package using pip:
	pip2 install pywin32 //Python 2.*
	pip3 install pywin32 //Python 3.*