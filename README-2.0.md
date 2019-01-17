# czy-nexus-commons-utils
   本库是发布到 [search.maven](https://search.maven.org/)  、 [mvnrepository](https://mvnrepository.com/)公共仓库的管理库。        
   (csdn教程博客)[https://blog.csdn.net/JavaWebRookie/article/details/80843653]、可通过maven方式下载源码查看注释。                
   (github工具类集库)[https://github.com/andyczy/czy-study-java-commons-utils]       
   (开源中国)[https://www.oschina.net/p/java-excel-utils]         
   
   
   2.0.1 版本：【推荐使用最新版本】       
          
         <!--
            maven：https://mvnrepository.com/artifact/com.github.andyczy/java-excel-utils
            教程文档：https://github.com/andyczy/czy-nexus-commons-utils/blob/master/README-3.2.md
         -->
        <dependency>        
            <groupId>com.github.andyczy</groupId>       
            <artifactId>java-excel-utils</artifactId>       
            <version>2.0.1</version>      
        </dependency> 
        
   
  [版本-2.0之前教程](https://github.com/andyczy/czy-nexus-commons-utils/blob/master/README-2.0.md)   
        
  [版本-3.2教程](https://github.com/andyczy/czy-nexus-commons-utils/blob/master/README-3.2.md)   
  
       
## 更新日志【2.0之前版本】       
###  导出使用函数: ExcelUtils.exportForExcel(......)
        * 可提供模板下载           
        * 自定义下拉列表：对每个单元格自定义下拉列表         
        * 自定义列宽：对每个单元格自定义列宽         
        * 自定义样式：对每个单元格自定义样式  
        * 自定义样式：单元格自定义某一列或者某一行样式            
        * 自定义单元格合并：对每个单元格合并 
        * 自定义：每个表格的大标题          
        * 自定义：对每个单元格固定表头    
        
        
### 导入配置使用函数： ExcelUtils.importForExcelData(......)
        * 获取多单元数据         
        * 自定义：多单元从第几行开始获取数据            
        * 自定义：多单元根据那些列为空来忽略行数据         

  
        
### 数据格式
   [javadoc 文档](https://oss.sonatype.org/service/local/repositories/releases/archive/com/github/andyczy/java-excel-utils/3.2/java-excel-utils-3.2-javadoc.jar/!/com/github/andyczy/java/excel/ExcelUtils.html)

   
   1、导出数据：参数 dataLists
   
        @Override
           public List<List<String[]>> exportBill(String deviceNo,String snExt,Integer parentInstId,String startDate, String endDate){
               List<List<String[]>> dataLists = new ArrayList<>();
               List<String[]> stringList = new ArrayList<>();
               PageInfo<BillInfo> pagePageInfo = getBillPage(1,10000,null,snExt,deviceNo,parentInstId,startDate,endDate);
               String[] valueString = null;
 
               String[] headers = {"序号","标题一","标题一","标题二","标题三","标题四","标题五","标题六"};
               String[] headersTwo = {" ","标题一小标题（合并用）","标题一小标题（合并用）"," "," "," "," "};
               stringList.add(headers);
               stringList.add(headersTwo);
                
               for (int i = 0; i < pagePageInfo.getList().size(); i++) {
                   valueString = new String[]{(i + 1) + "", pagePageInfo.getList().get(i).getSnExt(),
                           getNeededDateStyle(pagePageInfo.getList().get(i).getPayTime(),"yyyy-MM-dd hh:mm:ss"),
                           pagePageInfo.getList().get(i).getInstName(),pagePageInfo.getList().get(i).getStatisticsPrice()+"",
                           pagePageInfo.getList().get(i).getDeviceNo(),
                           pagePageInfo.getList().get(i).getWarning()==1?"是":"否"};
                   stringList.add(valueString);
               }
               listArray.add(stringList);
               return dataLists;
           }       
   
   2、自定义列宽：参数 columnMap
   
       参数说明：
       HashMap<Integer, HashMap<Integer, Integer>> columnMap = new HashMap<>();
       HashMap<Integer, Integer> mapColumn = new HashMap<>();
       //第一列、宽度为 3[3的大小就是两个12号字体刚刚好的列宽]（注意：excel从零行开始数）
       mapColumn.put(0, 3);  
       mapColumn.put(1, 20);
       mapColumn.put(2, 15);
       //第一个单元格列宽
       columnMap.put(1, mapColumn);
       
   3、自定义固定表头：参数 paneMap
   
       参数说明：
       HashMap paneMap = new HashMap();
       //第一个表格、第一行开始固定表头
       paneMap.put(1, 1); 
       
   
   4、自定义合并单元格：参数 regionMap
   
        参数说明：
        List<List<Integer[]>> regionMap = new ArrayList<>();
        List<Integer[]> regionList = new ArrayList<>();                  
        //代表起始行号，终止行号， 起始列号，终止列号进行合并。（注意：excel从零行开始数）
        regionList.add(new Integer[]{1, 1, 0, 10});
        regionList.add(new Integer[]{2, 3, 1, 1});
        //第一个表格设置。
        regionMap.put(1, regionList);
                                      
        
   5、自定义每个表格第几行或者是第几列的样式：参数 rowStyles / columnStyles
           
        参数说明：
        HashMap columnStyles = new HashMap();
        List list = new ArrayList();
        //1、样式（是否居中？，是否右对齐？，是否左对齐？， 是否加粗？，是否有边框？ ）
        list.add(new Boolean[]{true, false, false, false, true}); 
        //2、第几行或者是第几列（注意：excel从零行开始数）       
        list.add(new Integer[]{1, 3});   
        //3、颜色值（8是黑色、10红色等） 、颜色、字体、行高？（可不设置）                                        
        list.add(new Integer[]{10,14,null});    
        //第一表格                                 
        columnStyles.put(1,list);                                                     
        
   6、自定义每一个单元格样式：参数 styles
        
       参数说明：
       HashMap styles = new HashMap();
       List< List<Object[]>> stylesList = new ArrayList<>();
       List<Object[]> stylesObj = new ArrayList<>();
       List<Object[]> stylesObjTwo = new ArrayList<>();
       
       //1、样式一（是否居中？，是否右对齐？，是否左对齐？， 是否加粗？，是否有边框？ ）
       stylesObj.add(new Boolean[]{true, false, false, false, true});      
       //1、颜色值（8是黑色、10红色等） 、颜色、字体、行高？（可不设置）（必须放第二）
       stylesObj.add(new Integer[]{10, 12});                             
       //1、第五行、第一列（注意：excel从一开始算）
       stylesObj.add(new Integer[]{5, 1});                                  
       stylesObj.add(new Integer[]{6, 1});                                
       
       //2、样式二（必须放第一）
       stylesObjTwo.add(new Boolean[]{false, false, false, true, true}); 
       //2、颜色值（8是黑色、10红色等） 、颜色、字体、行高？（可不设置）（必须放第二）  
       stylesObjTwo.add(new Integer[]{10, 12,null});    
       //2、第二行第一列（注意：excel从一开始算）                 
       stylesObjTwo.add(new Integer[]{2, 1});                              
       
       stylesList.add(stylesObj);
       stylesList.add(stylesObjTwo);
       //第一个表格所有自定义单元格样式 
       styles.put(1, stylesList);                                             
             
   
   7、自定义忽略边框：参数 notBorderMap
   
       HashMap notBorderMap = new HashMap();
       //忽略边框（1行、5行）、默认是数据（除大标题外）是全部加边框的。
       notBorderMap.put(1, new Integer[]{1, 5});   
   
   
   8、自定义下拉列表值：参数 dropDownMap
      
       参数说明：
       HashMap dropDownMap = new HashMap();
       List<String[]> dropList = new ArrayList<>();                  
       //必须放第一：设置下拉列表的列（excel从零行开始数）
       String[] sheetDropData = new String[]{"1", "2", "4"};
       //下拉的值放在 sheetDropData 后面。        
       String[] sex = {"男,女"};                                   
       dropList.add(sheetDropData);
       dropList.add(sex);
       //第一个表格设置。
       dropDownMap.put(1,dropList); 
   
   9、导入配置：
        
       @param indexMap 多单元从第几行开始获取数据，默认从第二行开始获取（可为空，如 hashMapIndex.put(1,3); 第一个表格从第三行开始获取）
       @param continueRowMap 多单元根据那些列为空来忽略行数据（可为空，如 mapContinueRow.put(1,new Integer[]{1, 3}); 第一个表格从1、3列为空就忽略）
       
       
                   
### License
java-excel-utils is Open Source software released under the Apache 2.0 license.     