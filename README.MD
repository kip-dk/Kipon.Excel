# Library for streaming .NET object data between .NET and Excel

## Version alfa 1.
The very first alfa version of this library is taken form.

The aim for the project is to make it extreamly easy to take any tabular object data in .NET, stream these data to Excel, allow simple editing, and then get the data
back into .NET code.

Below example on streaming a simply object of type MyObjectData to excel:

(a NuGet package for .NET 4.7.2 on its way)

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


Last updated: 2019-08-10.