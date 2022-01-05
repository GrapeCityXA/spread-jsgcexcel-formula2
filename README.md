## [双剑合璧] SpreadJ+GcExcel 应对大量公式计算场景的解决方案（上）

Excel是我们日常必备的办公软件，用于处理日常各种表格事务。但很少有人知道，这个名字其实源自英语中的“Excellence”一词，代表着：卓越和优秀。而最能体现“卓越和优秀“的地方，就是Excel的公式计算功能，公式计算使得本身静态的表格变得“活”了起来。

基于Excel强大的公式计算能力，使得Excel已经在很多职场领域成为了精英们高效工作必备的神器。鲁迅曾经说过：“excel多学一公式，晚上多睡一小时。”

![输入图片说明](https://gcdn.grapecity.com.cn/data/attachment/forum/202108/26/110903zjoaqomsq7a7s7vf.png "在这里输入图片标题")


SpreadJS众所周知是一款能够在线运行的类Excel控件，它将Excel的功能搬到了线上的网页中，作为万能的程序猿，我们可以利用它将Excel的功能移植到我们的业务系统中，为我们的业务系统赋予Excel的能力。这就好比是关羽配上了赤兔马，如虎添翼。

![输入图片说明](https://gcdn.grapecity.com.cn/data/attachment/forum/202108/26/112550x7jelb0ge0fcvdld.png "在这里输入图片标题")

但是SpreadJS基于纯前端的设计使得在使用的过程中可能会遭遇到前端的性能瓶颈。前端资源是有限的，如果我们去加载一个包含大量公式计算的Excel，举个例子诸如地产行业的投资模型，金融保险行业的金融精算表格，财税行业的底稿等等。这些Excel中，公式个数一般在10W~20W这样的数量级上，并且其中还会包含大量复杂逻辑，嵌套的公式计算。这种情况下浏览器本身就会有所限制，已经不是SpreadJS力所能及的事情了。在用户体验上，就会表现为页面运行缓慢，甚至崩溃的情况。 

![输入图片说明](https://gcdn.grapecity.com.cn/data/attachment/forum/202108/27/183029pzwjllha886626x6.png "在这里输入图片标题")

当然，我们也不必惊慌失措，利用另一款组件GcExcel在服务端和性能的优势，与SpreadJS双剑合璧，可以有效的对上述场景进行优化。
下面就给大家来讲解专门应对此类问题的解决方案。



首先，我们先讲解思路：
1、利用GcExcel的服务端的特性和性能优势。在服务端对整个Excel进行加载和总体计算。
2、前端SpreadJS,只用负责页面结果的展示和与用户的操作交互。

架构图如下所示：

![输入图片说明](https://gcdn.grapecity.com.cn/data/attachment/forum/202108/27/185023btqkqwsx83ssu0fu.png "在这里输入图片标题")


根据这样的设计GcExcel可以有效分担原本SpreadJS的部分任务（这部分任务本身会大量的消耗前端性能），减轻前端压力。结构上避免了头重脚轻，更加匀称。

一些核心的代码实现：
1、GcExcel后端读取整个xlsx文件，第一次读取由于Excel文件之前会自己重算，所以可以巧妙利用这一点在读取后关闭后端计算引擎，这样可以节省一次重计算的时间。读取文件的时候尽量用流进行读取，流的效率要比字符串高效很多。

workbook.open(SpreadController.class.getClassLoader().getResourceAsStream("Excel/Wicked.xlsx"),options);
workbook.setEnableCalculation(false);

2、之后查询获取所有的Sheet工作表个数及名称及工作表显示隐藏的状态，这个后面有用


List<Sheet> sheetNameList = new ArrayList<Sheet>();
for(int i=0;i<workbook.getWorksheets().getCount();i++) {
  Sheet sheet = new Sheet();
  IWorksheet worksheet = workbook.getWorksheets().get(i);
  sheet.setName(worksheet.getName());
  if(Visibility.Hidden.equals(worksheet.getVisible())) {
    sheet.setVisiable(false);
  }else {
    sheet.setVisiable(true);
  }
  sheetNameList.add(sheet);
}



3、获取activeSheet的名称和ssjson，将上述整个合并到结果中返回，这里可以将返回结果进行压缩，进一步返回大小，这样可有效加快网络请求的时间（事例中用了GZip的通用压缩工具）。


IWorksheet activeSheet = workbook.getActiveSheet();
String activeSheetName = activeSheet.getName();
returnMap.put("activeSheetName", activeSheetName);

String result = activeSheet.toJson();
result = GZip.compress(result);
returnMap.put("sheetJSON", result);




4.前端SpreadJS接收到结果后，根据第二步工作表的名称及个数新建等量的工作表并改名改状态。这里完了之后可以将前端计算引擎挂起（本方案无需前端计算，但需要有前端的公式显示作为提示）。然后获取activeSheet并反序列化第三步生成的ssjson



spread.setSheetCount(length);
for(var i=0;i<length;i++){
  spread.getSheet(i).name(data.sheetNames[i].name);
  if(data.sheetNames[i].visiable == false){
    spread.getSheet(i).visible(false);
  }
  if(data.sheetNames[i].name == data.activeSheetName){
    spread.setActiveSheetIndex(i);
  }
}
spread.suspendCalcService(false);
var activeSheet = spread.getActiveSheet();
var json = data.sheetJSON;
json = ungzipString(json);
json = JSON.parse(json);
activeSheet.fromJSON(json);



这样初步加载就完成了，还没有睡着的小伙伴可以继续接着向下看。


4. 根据刚刚的设计，前端SpreadJS中，目前只有activeSheet是有货的，其余都是空的sheet。这个时候我们就需要去监听事件，通过事件驱动来进行缓式加载。这里监听了ActiveSheetChanging事件，用于在sheet切换的时机去后端获取新的activeSheet的ssjson。这里以防用户在sheet上进行修改，同时获取了脏数据，一并传给后端。最终将后端的返回结果在前端反序列化。


spread.bind(GC.Spread.Sheets.Events.ActiveSheetChanging, function (sender, args) {
                var oldSheet = args.oldSheet;
                var dirtyCells = oldSheet.getDirtyCells();
                var newSheet = args.newSheet;
                var sheetChange = {
  oldSheetName:oldSheet.name(),
  newSheetName:newSheet.name(),
  dirtyCells:dirtyCells        
                }
                $.ajax({
  url: "getSheet",
  type:"POST",
  data:JSON.stringify(sheetChange),
  contentType: 'application/json',
  async:false,
  success: function (data) {
    if (data != null) {
      var json = ungzipString(data);
      json = JSON.parse(json);
      newSheet.fromJSON(json);
    }
  }
                    });
            });



5.后端拿到上述信息同步到整体的workbook中，然后重新计算得到计算后的结果，在将新的activeSheet序列化成ssjson返回


IWorksheet oldSheet = workbook.getWorksheets().get(sheetChange.getOldSheetName());
                if(sheetChange.getDirtyCells().size()>0) {
  for(int i=0;i<sheetChange.getDirtyCells().size();i++) {
    Map<String,Object> dirtyCell = sheetChange.getDirtyCells().get(i);
    int row = (int) dirtyCell.get("row");
    int col = (int) dirtyCell.get("col");
    oldSheet.getRange(row, col).setValue(dirtyCell.get("newValue"));
  }
  workbook.setEnableCalculation(true);
                }
               
                IWorksheet newSheet = workbook.getWorksheets().get(sheetChange.getNewSheetName());
                String result = null;
                if(newSheet!=null) {
  result = newSheet.toJson();
  result = GZip.compress(result);        
                }
                return result;

