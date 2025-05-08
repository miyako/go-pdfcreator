# go-pdfcreator

PDFCreator has several [editions](https://www.pdfforge.org/pdfcreator/editions).

The free edtion indeed includes COM-interface:

<img src="https://github.com/user-attachments/assets/dda0a09c-ad8a-4f6f-98a3-750016a3644a" width=500 height=auto /> 

However, the necessary DLL (`PDFCreator.COM.dll`) is NOT installed with the [free installer](https://www.pdfforge.org/pdfcreator/download).

Instead, you can download "PDFCreator Professional" setup and "Request Trial".

<img src="https://github.com/user-attachments/assets/8c77300d-63f9-488a-8af7-2fdd66940827" width=500 height=auto /> 

A trial license key is sent to your email and allows installation.

But first, make sure you install [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/en-us/download/dotnet/8.0)

<img src="https://github.com/user-attachments/assets/94f9c528-c4e3-4b29-9117-a121f74c70b5" width=500 height=auto /> 

> [!NOTE]
> Install the Intel (x64) version, even if Windows is ARM

<img src="https://github.com/user-attachments/assets/38c41ef1-b8df-4cf8-81b1-d2253590c88b" width=500 height=auto />  

If you have

* PDFCreator 
* `PDFCreator.COM.dll`
* .NET 8 Desktop Runtime (x64)

installed, you can register the [COM interface](https://docs.pdfforge.org/pdfcreator/en/pdfcreator/com-interface/#) with the following CLI

```
SetupHelper.exe /ComInterface=Register
```

<img src="https://github.com/user-attachments/assets/654f673e-3614-40fe-8852-8e1af40a64d3" width=500 height=auto />  
