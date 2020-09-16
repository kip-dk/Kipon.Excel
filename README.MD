# Library for streaming .NET object data between .NET and Excel

## Version alfa 1.
The very first alfa version of this library is taken form.

The aim for the project is to make it extreamly easy to take any tabular object data in .NET, stream these data to Excel, allow simple editing, and then get the data
back into .NET code.


### Installation (NuGet)

Published via NuGet: 

Kipon.Excel(https://www.nuget.org/packages/Kipon.Excel/)

```console
PM> Install-Package Kipon.Excel -Version 1.0.14-alpha
```

### Simple usage

Below example on streaming a simply object of type MyObjectData to excel:


```csharp
using Kipon.Excel.Linq;

namespace MyExcelApp 
{
    public static class Program 
	{ 
		MyObjetData[] myDatas = MethodToCollectionSomeData();

		byte[] excelData = myDatas.toExcel();
		System.IO.File.WriteAllBytes("myfile.xlsx", excelData);
	}
}

```

Below example of getting the data back into .NET again

```csharp
using Kipon.Excel.Linq;

namespace MyExcelApp 
{
	public static class Program
	{
		using (var file = new System.IO.FileStream("myfile.xlsx")) 
		{
		    MyObjetData[] myDatas = file.ToArray<MyObjectData>();
		}
	}
}
```



Last updated: 2020-09-16.