# Library for streaming .NET object data between .NET and Excel, 2022-08-07

## C# to Excel, with a streaming approach.
This library aim to add simplicity in the process of taging C# data, and put them into editable excel sheets, and the get the data back into a C# context without to much of an effort.

The streaming should be as simple as using any other sereializer.

The library is build on top of DocumentFormat.OpenXml. This library allow you to interact with excel data, but it is a very raw impl. the simply maps the XML structure of an excel sheet into a 
rather generic C# model. It is hard to work with, and full of strange pitfalls.

This library tries to fit that gab. 

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


### Changelog

#### 1.0.24
Fixed decimals format. If a property as marked with Decimals attribute, the content of the cell will be number styled accordingly.  0 .. 5 decimals is supported. Beyond 5 decimals, default number format will be used.

### 1.0.21 First official release

#### 1.0.19-alpha

Support write of Dictionary properties, A dictionary property must be ether defined with ColumnAttribute (IsIndex = true) or decorated with IndexColumnAttribute



Last updated: 2021-01-24.