import java.util.List;
import java.util.Map;

/**
 * Created by zhouchunjie on 16/1/22.
 */
public class Main {

    public final static void main(String[] args) {

        String path = "src/main/resources/sample.xlsx";
        try {
            List<Map<String, String>> result = TransformExcelUtil.readXls(path);
            System.out.print(result);
        } catch (Exception e) {
            System.out.print("读取失败");
        }



    }

}
