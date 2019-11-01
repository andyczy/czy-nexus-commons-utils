# czy-nexus-commons-utils
   (本库)[https://github.com/andyczy/czy-nexus-commons-utils]是发布到 [search.maven](https://search.maven.org/)  、 [mvnrepository](https://mvnrepository.com/)公共仓库的管理库。    
   (csdn教程博客)[https://blog.csdn.net/JavaWebRookie/article/details/80843653]、可通过maven方式下载源码查看注释。                
   (github工具类集库)[https://github.com/andyczy/czy-study-java-commons-utils]    
   (开源中国)[https://www.oschina.net/p/java-excel-utils]          
   
   推荐使用最新版本：        
          
         <!--
            maven：https://mvnrepository.com/artifact/com.github.andyczy/java-excel-utils
            教程文档：https://github.com/andyczy/czy-nexus-commons-utils/blob/master/README-3.2.md
         -->
        <dependency>        
            <groupId>com.github.andyczy</groupId>       
            <artifactId>java-excel-utils</artifactId>       
            <version>4.0</version>      
        </dependency> 
 
  [教程说明](https://github.com/andyczy/czy-nexus-commons-utils/blob/master/README-Andyczy.md)   
  [本地输出测试](https://github.com/andyczy/czy-nexus-commons-utils/blob/master/README-Local-Test.md)   
   
  亲自测试：WPS、office 07、08、09、10、11、12、16 能正常打开。其他版本待测试！               
  注:POI SXSSFWorkbook 最高限制1048576行,16384列           

### 功能说明
    单表百万数据量导出时样式设置过多，导致速度慢（行、列、单元格样式暂时控制10万行、超过无样式）                          
    大数据量情况下一般不会每个单元格设置样式、不然很难解决内存溢出等问题。                 
    修改输出流（只能输出一次、如 response 响应输出，则不会输出到本地路径的。）                                   
    修改注释                            
    新增函数【ExcelUtils.testLocalNoStyleNoResponse() 、本地测试：输出到本地路径】                  
    新增函数【ExcelUtils.exportForExcelsNoStyle()、无样式（行、列、单元格样式）推荐使用这个函数、提高速度】                
    初始化函数：ExcelUtils.setExcelUtils() 更改为 ExcelUtils.initialization()          
    属性：columnMap 更改为 setMapColumnWidth
    
    目前导出速度：
    （单表）1万行、20列：1.6秒            
    （单表）10万行、20列：11秒                 
    （单表）20万行、20列：27秒     
    （单表）104万行、20列：46秒            
    
    （4张表）1*4万行、20列：6秒           
    （4张表）10*4万行、20列：33秒                     
    （4张表）20*4万行、20列：61秒
    （4张表）100*4万行、20列：85秒
             
    【4.0】新增 LocalExcelUtils 对象、Test 本地测试、CommonsUtils工具类
         
            
### 最新日志（4.0版本没有、4.1没有上传到maven）                   
    1、是否添加边框改为是否忽略边框？默认单元格都带边框。
    2、添加导出图片。  
    3、可设置默认列宽大小。默认是16
    4、可设置默认字体大小。默认是12
    5、删除：导出函数 ExcelUtils.exportForExcel(......)过期、4.0以下版本有。
    
    
     
### 实现功能：
    1、自定义导入数据格式，支持配置时间、小数点类型（支持单/多sheet）              
    2、浏览器导出Excel文件、模板文件（支持单/多sheet）           
    3、指定路径生成Excel文件（支持单/多sheet）           
    4、自定义样式：行、列、某个单元格（字体大小、字体颜色、左右对齐、居中、是否忽略边框。支持单/多sheet）           
    5、自定义固定表头（支持单/多sheet）            
    6、自定义下拉列表值（支持单/多sheet）           
    7、自定义合并单元格、自定义列宽、自定义大标题（支持单/多sheet）
    8、导出图片、图片地址和数据一样，只要是能访问的图片都可以导出（有需求、图片大小待解决），图片格式：.JPEG|.jpeg|.JPG|.jpg|.png|.gif
           

# 感谢支持、感谢你们（排名不分先后）
蒙蒙的雨（3元微信）、阿星支付宝（100支付宝）、李凯（5元微信）、blue（5元微信2019-03-28）、鹏飞（50支付宝2019-06-05）、啊哈（3元微信19-06-26）、84644574*(QQ 4元19-07-08)                  
                  
![支持一下](https://github.com/andyczy/czy-nexus-commons-utils/blob/master/sqm.png)                        
        
       
   
### License
java-excel-utils is Open Source software released under the Apache 2.0 license.     