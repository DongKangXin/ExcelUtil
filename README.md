# ExcelUtil
Excel的文件导入解析成对象集合
基于阿帕奇的POI项目将EXCEL文件读取并且转化成对象集合

最新版本更新为注解编程模式

## 使用方法 ：

### 待转换的类
```
public class Student {

    //Header注解的value为对应的Excel文件中的表头
    @Header("姓名")
    private String name;

    @Header("年龄")
    private Integer age;

    @Header("分数")
    private Double score;

    @Header("出生日期")
    private Date date;
}
```

### 在程序中调用
```
public static void main(String[] args){
    Collection<Object> objects = ExcelUtil.parseExcel(new FileInputStream(file), Object.class);
}
