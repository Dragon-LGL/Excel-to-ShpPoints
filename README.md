# Excel-to-ShpPoints
## Excel文件转矢量点
* CSVtoPointShp.py为单文件处理程序
* ExcelToPointShp_Batch.py为批量处理程序

### 代码注释：  
```field_name = ogr.FieldDefn("Name", ogr.OFTString)```
设置矢量属性
  * ogr.OFTString 字符串
  * ogr.OFTReal 数值  

<br>可选参数：  
  * 0 = ogr.OFTInteger  
  * 1 = ogr.OFTIntegerList
  * 2 = ogr.OFTReal  
  * 3 = ogr.OFTRealList  
  * 4 = ogr.OFTString  
  * 5 = ogr.OFTStringList  
  * 6 = ogr.OFTWideString  
  * 7 = ogr.OFTWideStringList  
  * 8 = ogr.OFTBinary  
  * 9 = ogr.OFTDate  
  * 10 = ogr.OFTTime  
  * 11 = ogr.OFTDateTime  
