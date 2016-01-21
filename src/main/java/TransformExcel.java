import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by zhouchunjie on 16/1/21.
 */
public class TransformExcel {

    public final static void main(String[] args) {

        String path = "src/main/resources/sample.xlsx";
        try {
            List<Map<String, String>> result = readXls(path);
            System.out.print(result);
        } catch (Exception e) {
            System.out.print("读取失败");
        }
    }

    private static List<Map<String, String>> readXls(String path) throws Exception {
        //载入文件
        InputStream is = new FileInputStream(path);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);

        List<Map<String, String>> result = new ArrayList<Map<String, String>>();
        //循环sheet页面
        for (XSSFSheet xssfSheets : xssfWorkbook) {
            if (xssfSheets == null) {
                continue;
            }
            //获取标题行
            XSSFRow firstRow = xssfSheets.getRow(0);
            //开始遍历,获取内容
            for (int rowNum = 1; rowNum <= xssfSheets.getLastRowNum(); rowNum++) {
                //获取当前行
                XSSFRow xssfRow = xssfSheets.getRow(rowNum);
                //获取行最小index和最大index
                int minColIx = xssfRow.getFirstCellNum();
                int maxColIx = xssfRow.getLastCellNum();
                Map<String, String> rowMap = new HashMap<String, String>();

                //遍历row,获取cell
                for (int colIx = minColIx; colIx < maxColIx; colIx++) {
                    //获取当前单元格所对应标题
                    XSSFCell titleCell = firstRow.getCell(colIx);
                    //获取当前单元格
                    XSSFCell cell = xssfRow.getCell(colIx);
                    if (cell == null) {
                        //单元格内容为空,添加空字符串
                        rowMap.put(titleCell.getStringCellValue(), "");
                        continue;
                    }
                    //单元格内容不为空,添加字段对应内容
                    rowMap.put(titleCell.getStringCellValue(), getStringVal(cell));
                }

                result.add(rowMap);
            }
        }
        return result;
    }

    /**
     * 单元格格式转化
     *
     * @param cell
     * @return
     */
    private static String getStringVal(XSSFCell cell) {
        switch (cell.getCellType()) {
            case XSSFCell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue() ? "TRUE" : "FALSE";
            case XSSFCell.CELL_TYPE_FORMULA:
                return cell.getCellFormula();
            case XSSFCell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    //数字类型,且符合日期格式,转换成日期
                    DateFormat sdf = new SimpleDateFormat("dd-MM-yyyy HH:mm");
                    return sdf.format(cell.getDateCellValue());
                }
                //非日期格式,转换成String
                //TODO 如果值为浮点型,可能会使数字失真
                cell.setCellType(cell.CELL_TYPE_STRING);
                return cell.getStringCellValue();
            case XSSFCell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            default:
                //未知类型转换为空字符串
                return "";
        }
    }
}
