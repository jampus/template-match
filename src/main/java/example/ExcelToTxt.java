package example;


import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.ReflectUtil;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.serializer.SerializerFeature;
import com.spire.xls.FileFormat;
import com.spire.xls.Workbook;
import example.dto.ReportDTO;

import java.lang.reflect.Field;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.Map;


public class ExcelToTxt {
    public static void main(String[] args) {

        // 获取 resources 目录的 URL
        String basePath = ExcelToTxt.class.getClassLoader().getResource("").getPath().replace("target/classes/","");

        String excelTemplatePath = basePath + "excel/监测报告_template.xlsx";
        String excelPath = basePath + "excel/监测报告.xlsx";

        //临时文件
        String csvTemplatePath = basePath + "excel/监测报告_template.csv";
        String csvPath = basePath + "excel/监测报告.csv";

        Workbook workbookTemplate = new Workbook();
        workbookTemplate.loadFromFile(excelTemplatePath);
        workbookTemplate.saveToFile(csvTemplatePath, FileFormat.CSV);

        Workbook workbook = new Workbook();
        workbook.loadFromFile(excelPath);
        workbook.saveToFile(csvPath, FileFormat.CSV);

        //读取文件内容
        final String templateContent = FileUtil.readString(csvTemplatePath, StandardCharsets.UTF_8);
        final String content = FileUtil.readString(csvPath, StandardCharsets.UTF_8);
        FileUtil.del(csvTemplatePath);
        FileUtil.del(csvPath);


        //匹配差异
        diff_match_patch dmp = new diff_match_patch();
        final LinkedList<diff_match_patch.Diff> diffs = dmp.diff_main(templateContent, content);
        dmp.diff_cleanupEfficiency(diffs);

        //收集变量名和值
        final HashMap<String, Object> variables = new HashMap<>();
        diff_match_patch.Diff leftDiff = null;
        for (diff_match_patch.Diff diff : diffs) {
            if (diff.operation == diff_match_patch.Operation.EQUAL) {
                continue;
            }
            if (diff.operation == diff_match_patch.Operation.INSERT || diff.operation == diff_match_patch.Operation.DELETE) {
                if (leftDiff == null) {
                    leftDiff = diff;
                    continue;
                }
                if (leftDiff.operation == diff.operation) {
                    leftDiff = diff;
                } else {
                    String key = leftDiff.text.replaceAll("\"", "");
                    String value = diff.text.replaceAll("\"", "");
                    variables.put(key, value);
                    leftDiff = null;
                }
            }
        }
        //转成对象
        final Object o = mapToObject(variables, ReportDTO.class, "");
        System.out.println(JSON.toJSONString(o, SerializerFeature.PrettyFormat));

    }

    public static Object mapToObject(Map<String, Object> variables, Class clazz, String parent) {
        final Object newInstance = ReflectUtil.newInstance(clazz);
        final Field[] fields = ReflectUtil.getFields(clazz);
        for (Field field : fields) {
            if (field.getType().isAssignableFrom(String.class)) {
                final String fieldName = parent + field.getName();
                ReflectUtil.setFieldValue(newInstance, field, variables.get(fieldName));
            } else {
                final Object object = mapToObject(variables, field.getType(), parent + field.getName() + "_");
                ReflectUtil.setFieldValue(newInstance, field, object);
            }
        }
        return newInstance;
    }
}
