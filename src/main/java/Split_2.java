import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Pattern;

public class Split_2 {
    public static void main(String[] args) {
        try {
            // 打开源文件
            FileInputStream file = new FileInputStream(new File("data/software.xlsx"));
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);

            // 输出文件
            Workbook workbook1 = new XSSFWorkbook();
            Sheet sheet1 = workbook1.createSheet("Sheet1");
            Row out =sheet1.createRow(0);
            for (String keyword : index.keySet()) {
                Cell cell = out.createCell(index.get(keyword));
                cell.setCellValue(keyword);
            }

            // 当前的行数以及姓名
            int num = 1;
            String current = "";

            for (Row row : sheet) {
                // 获取姓名
                String name = row.getCell(0).getStringCellValue();
                if (!name.equals(current)) {
                    // 姓名变化 -> 修改名字和行数
                    current = name;
                    out = sheet1.createRow(num++);
                    Cell cell = out.createCell(0);
                    cell.setCellValue(current);
                    System.out.println(current);
                }

                // 获取数据
                Cell data = row.getCell(1);
                if (null != data) {
                    String content = data.getStringCellValue();
                    if (independence(content)) continue;

                    // 根据关键字将文本进行划分
                    String patternString = "(?=(" + String.join("|", index1.keySet()) + "\\n))";
                    Pattern pattern = Pattern.compile(patternString);
                    String[] parts = pattern.split(content);

                    for (String part : parts) {
                        // 再将文本块划分为关键字和内容
                        int firstNewLineIndex = part.indexOf("\n");
                        if (-1 == firstNewLineIndex) continue;
                        String firstLine =part.substring(0, firstNewLineIndex);
                        String remainingLines = part.substring(firstNewLineIndex + 1);

                        if (null == index.get(firstLine)) {
                            // 若关键字为”个人信息“ -> 在将个人信息里面的内容进行划分
                            String[] lines = remainingLines.split("\n");
                            if (independence(lines[0])) continue;

                            // 根据冒号划分
                            for (String line : lines) {
                                String[] messages = line.split("：", 2);
                                save(out, messages[0], messages[1]);
                            }
                        } else {
                            // 其他关键字内容可以直接保存
                            save(out, firstLine, remainingLines);
                        }

                    }
                }
            }
            FileOutputStream fos = new FileOutputStream("data/software_out.xlsx");
            workbook1.write(fos);
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // 排除干扰信息
    public static boolean independence(String content) {
        return content.equals("相关教师") || content.equals("导航") || content.equals("") || content.equals("访问");
    }

    public static void save(Row out, String firstLine, String remainingLines) {
        if (null == index.get(firstLine) || null != index.get(remainingLines)) return;
        Cell cell = out.createCell(index.get(firstLine));
        cell.setCellValue(remainingLines);
    }

    public static final Map<String, Integer> index = new HashMap<>() {
        {
            put("姓名", 0);
            put("部门", 1);
            put("性别", 2);
            put("专业技术职务", 3);
            put("毕业院校", 4);
            put("学位", 5);
            put("学历", 6);
            put("邮编", 7);
            put("联系电话", 8);
            put("传真", 9);
            put("电子邮箱", 10);
            put("办公地址", 11);
            put("通讯地址", 12);
            put("教育经历", 13);
            put("工作经历", 14);
            put("个人简介", 15);
            put("社会兼职", 16);
            put("研究方向", 17);
            put("研究领域", 17);
            put("开授课程", 18);
            put("科研项目", 19);
            put("学术成果", 20);
            put("荣誉及奖励", 21);
            put("招生与培养", 22);
        }
    };

    public static final Map<String, Integer> index1 = new HashMap<>() {
        {
            put("个人资料\n", 0);
            put("教育经历\n", 1);
            put("工作经历\n", 2);
            put("个人简介\n", 3);
            put("社会兼职\n", 4);
            put("研究方向\n", 5);
            put("研究领域\n", 5);
            put("开授课程\n", 6);
            put("科研项目\n", 7);
            put("学术成果\n", 8);
            put("荣誉及奖励\n", 9);
            put("招生与培养\n", 10);
        }
    };
}
