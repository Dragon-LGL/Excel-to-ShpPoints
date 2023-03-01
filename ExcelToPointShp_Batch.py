#-*- coding : utf-8-*-
"""
# —*— ProjectName：ExcelToPointShp_Batch
# —*— Author：李国龙
# —*— CreateTime：2022年12月13日13:28:31
# —*— Vision：1.0.0
# —*— Information:批量Excel文件转点矢量
"""
from osgeo import ogr ,osr
import csv
import pandas as pd
import os


def xlsx_to_csv_pd(xlsx_file):
    data_xls = pd.read_excel(xlsx_file, index_col=0)
    csv_file = xlsx_file.replace('.xlsx', '.csv')
    data_xls.to_csv(csv_file, encoding='gbk')
    return csv_file

def xls_to_csv_pd(xls_file):
    xls_file = pda.read_excel(xls_file)
    csv_file = xlsx_file.replace('.xls', '.csv')
    xls_file.to_csv(csv_file, encoding="gbk")
    return csv_file

def Create_PointShp(Longitudes, Latitudes, shpfile):
    ## 生成点矢量文件 ##
    driver = ogr.GetDriverByName("ESRI Shapefile")
    data_source = driver.CreateDataSource(shpfile)   ## shp文件名称
    srs = osr.SpatialReference()
    srs.ImportFromEPSG(32650) ## 32650表示坐标系WGS_1984_UTM_Zone_50N
    layer = data_source.CreateLayer("Point", srs, ogr.wkbPoint) ## 图层名称要与shp名称一致
    field_name = ogr.FieldDefn("Name", ogr.OFTString) ## 设置属性
    field_name.SetWidth(20)  ## 设置长度
    layer.CreateField(field_name)  ## 创建字段
    field_Longitude = ogr.FieldDefn("Longitude", ogr.OFTReal)  ## 设置属性
    layer.CreateField(field_Longitude)  ## 创建字段
    field_Latitude = ogr.FieldDefn("Latitude", ogr.OFTReal)  ## 设置属性
    layer.CreateField(field_Latitude)  ## 创建字段
    feature = ogr.Feature(layer.GetLayerDefn())


    for i in range(len(Longitudes)):
        AddPoint(Longitudes[i], Latitudes[i], feature, layer)
        # print('正在输入第{}个点······'.format(i+1))
    feature = None ## 关闭属性
    data_source = None ## 关闭数据

def AddPoint(Longitude, Latitude, feature, layer):
    feature.SetField("Name", "point")  ## 设置字段值
    feature.SetField("Longitude", str(Longitude))  ## 设置字段值
    feature.SetField("Latitude", str(Latitude))  ## 设置字段值
    wkt = "POINT(%f %f)" % (float(Longitude), float(Latitude)) ## 创建点
    point = ogr.CreateGeometryFromWkt(wkt) ## 生成点
    feature.SetGeometry(point)  ## 设置点
    layer.CreateFeature(feature)  ## 添加点


def CSVToShape(csvfile, shpfile):
    #	CSV的读取：
    print("Running:Open CSV File")
    Lon = []
    Lat = []
    FID = []
    with open(csvfile, 'r', encoding='UTF-8') as f:  # 打开CSV文件
        reader = csv.reader(f)
        # Lon_data = list(reader)
        for i in reader:
            Lon.append(i[2])  # 经度所在的列下标
            Lat.append(i[3])  # 纬度所在的列下标
            FID.append(i[0])  # 属性数据
        f.close()
    #	读取高程、经度、纬度
    print("Running:Reading Data")
    Create_PointShp(Lon[1:], Lat[1:], shpfile)

if __name__ == '__main__':

    excel_file_dir = 'E:/test' #输入的Excel文件路径
    shp_file = 'E:/test/point'   #输出的Shp文件路径
    files = os.listdir(excel_file_dir)

    for file in files:
        print(file)
        if file.find('.csv') == -1:
            if file.find('.xlsx') >= 0:
                xlsx_to_csv_pd(file)
                print(file + "文件已转为CSV文件！")
            elif file.find('.xls') >= 0:
                xls_to_csv_pd(file)
                print(file + "xls文件已转为CSV文件！")
            else:
                print(file + '未转换！')
        CSVToShape(file, shp_file)
    print("Output ShapeFile Success!")