package com.leon.hello.excel.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.util.ListUtils;
import com.alibaba.excel.util.MapUtils;
import com.alibaba.fastjson.JSON;
import com.leon.hello.excel.entity.DownloadData;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.*;

/**
 * @PROJECT_NAME: hello-excel
 * @CLASS_NAME: EasyexcelDemoController
 * @AUTHOR: OceanLeonAI
 * @CREATED_DATE: 2021/11/22 14:32
 * @Version 1.0
 * @DESCRIPTION:
 **/
@RestController
public class EasyexcelDemoController {

    /**
     * 文件下载（失败了会返回一个有部分数据的Excel）
     * <p>
     * 1. 创建excel对应的实体对象 参照{@link DownloadData}
     * <p>
     * 2. 设置返回的 参数
     * <p>
     * 3. 直接写，这里注意，finish的时候会自动关闭OutputStream,当然你外面再关闭流问题不大
     */
    @GetMapping("download")
    public void download(HttpServletResponse response) throws IOException {
        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
        String fileName = URLEncoder.encode("测试", "UTF-8").replaceAll("\\+", "%20");
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");

        EasyExcel.write(response.getOutputStream())
                .head(getMockDataHead()) // 动态列名
                .sheet("这里是导出excel的sheet页名称") // sheet 页名称
                .doWrite(getMockDataList()); // 动态数据
    }

    @GetMapping("downloadExcel")
    public void downloadExcel(HttpServletResponse response) throws IOException {
        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        try {
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setCharacterEncoding("utf-8");
            // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
            String fileName = URLEncoder.encode("测试", "UTF-8").replaceAll("\\+", "%20");
            response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
            // 这里需要设置不关闭流
            EasyExcel.write(response.getOutputStream())
                    .head(getMockDataHead()) // 动态列名
                    .autoCloseStream(Boolean.FALSE)
                    .sheet("模板")
                    .doWrite(getMockDataList());
            // ============ 测试异常情况 begin ============
            // String str = null;
            // str.toString();
            // ============ 测试异常情况   end ============
        } catch (Exception e) {
            // 重置response
            response.reset();
            response.setContentType("application/json");
            response.setCharacterEncoding("utf-8");
            Map<String, String> map = MapUtils.newHashMap();
            map.put("status", "failure");
            map.put("message", "下载文件失败" + e.getMessage());
            response.getWriter().println(JSON.toJSONString(map));
        }
    }

    /**
     * 文件下载并且失败的时候返回json（默认失败了会返回一个有部分数据的Excel）
     *
     * @since 2.1.1
     */
    @GetMapping("downloadFailedUsingJson")
    public void downloadFailedUsingJson(HttpServletResponse response) throws IOException {
        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        try {
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setCharacterEncoding("utf-8");
            // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
            String fileName = URLEncoder.encode("测试", "UTF-8").replaceAll("\\+", "%20");
            response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
            // 这里需要设置不关闭流
            EasyExcel.write(response.getOutputStream(), DownloadData.class).autoCloseStream(Boolean.FALSE).sheet("模板")
                    .doWrite(getMockDataEntity());
            // ============ 测试异常情况 begin ============
            // String str = null;
            // str.toString();
            // ============ 测试异常情况   end ============
        } catch (Exception e) {
            // 重置response
            response.reset();
            response.setContentType("application/json");
            response.setCharacterEncoding("utf-8");
            Map<String, String> map = MapUtils.newHashMap();
            map.put("status", "failure");
            map.put("message", "下载文件失败" + e.getMessage());
            response.getWriter().println(JSON.toJSONString(map));
        }
    }

    /**
     * 模拟动态获取Excel列名
     *
     * @return
     */
    private List<List<String>> getMockDataHead() {
        List<List<String>> list = ListUtils.newArrayList();
        List<String> head0 = ListUtils.newArrayList();
        head0.add("姓名" + System.currentTimeMillis());
        List<String> head1 = ListUtils.newArrayList();
        head1.add("年龄" + System.currentTimeMillis());
        List<String> head2 = ListUtils.newArrayList();
        head2.add("地址" + System.currentTimeMillis());
        list.add(head0);
        list.add(head1);
        list.add(head2);
        return list;
    }

    /**
     * 模拟Map<String,Object>数据
     * key必须为列下标才有效，否则无法将数据写入excel
     *
     * @return
     */
    private List<Map<String, Object>> getMockDataMap() {
        List<Map<String, Object>> list = ListUtils.newArrayList();
        for (int i = 0; i < 10; i++) {
            Map<String, Object> map = new HashMap<>();
            map.put("1", "leon" + (i + 1));
            map.put("2", (i + 1));
            map.put("3", "成都市高新区孵化园" + (i + 1));
            list.add(map);
        }
        return list;
    }

    /**
     * 模拟Map<String,Object>数据
     * key必须为列下标才有效
     *
     * @return
     */
    private List<Map<Integer, Object>> getMockDataMapKeyIsIndex() {
        List<Map<Integer, Object>> list = ListUtils.newArrayList();
        for (int i = 0; i < 10; i++) {
            Map<Integer, Object> map = new HashMap<>();
            map.put(0, "leon" + (i + 1));
            map.put(1, (i + 1));
            map.put(2, "成都市高新区孵化园" + (i + 1));
            list.add(map);
        }
        return list;
    }

    /**
     * 模拟返回List数据
     * 数据顺序和列头对应
     *
     * @return
     */
    private List<List<Object>> getMockDataList() {
        List<List<Object>> list = ListUtils.newArrayList();
        for (int i = 0; i < 10; i++) {
            List<Object> objList = new ArrayList<>();
            objList.add("zhagnsan" + (i + 1));
            objList.add("age" + (i + 1));
            objList.add("address" + (i + 1));
            list.add(objList);
        }

        return list;
    }

    /**
     * 实体数据
     *
     * @return
     */
    private List<DownloadData> getMockDataEntity() {
        List<DownloadData> list = ListUtils.newArrayList();
        for (int i = 0; i < 10; i++) {
            DownloadData data = new DownloadData();
            data.setString("字符串" + 0);
            data.setDate(new Date());
            data.setDoubleData(0.56);
            list.add(data);
        }
        return list;
    }

}
