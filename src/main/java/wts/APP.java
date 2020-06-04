package wts;

import com.alibaba.fastjson.JSONObject;
import okhttp3.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;

public class APP {
    /*
    获取token
    */
    public static String getToken() throws IOException {
        OkHttpClient client = new OkHttpClient().newBuilder()
                .build();
        MediaType mediaType = MediaType.parse("application/json;charset=UTF-8,text/plain");
        RequestBody body = RequestBody.create(mediaType, "{\"user_account\":\"举报投诉中心\",\"password\":\"\",\"captcha\":\"\",\"terminal\":\"pc\",\"dynamic_code\":\"\",\"local\":\"zh-CN\",\"sms_token\":\"\",\"sms_verify_code\":\"\",\"login_set\":\"\"}");
        Request request = new Request.Builder()
                .url("http://120.221.150.148:8010/eoffice10/server/public/api/auth/login")
                .method("POST", body)
                .addHeader("Accept", "application/json, text/plain, */*")
                .addHeader("Accept-Encoding", "gzip, deflate")
                .addHeader("Accept-Language", "zh-CN,zh;q=0.9")
                .addHeader("Connection", "keep-alive")
                .addHeader("Content-Length", "165")
                .addHeader("Content-Type", "application/json;charset=UTF-8")
                .addHeader("Cookie", "io=H-toQ_Vi7yMnUWmSAEDo")
                .addHeader("Host", "120.221.150.148:8010")
                .addHeader("Origin", "http://120.221.150.148:8010")
                .addHeader("Referer", "http://120.221.150.148:8010/eoffice10/client/app/web/login.html")
                .addHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.25 Safari/537.36 Core/1.70.3756.400 QQBrowser/10.5.4039.400")
                .addHeader("Content-Type", "text/plain")
                .build();
        Response response = client.newCall(request).execute();
        String token = JSONObject.parseObject(JSONObject.parseObject(response.body().string()).getString("data")).getString("token");
        System.out.println("获取TOKEN成功！");
        return token;
    }

    /*
    获取run_id：通过新建流程来获取
    */
    public static String getRun_id(String token) throws IOException {
        OkHttpClient client = new OkHttpClient().newBuilder()
                .build();
        MediaType mediaType = MediaType.parse("application/json;charset=UTF-8");
        RequestBody body = RequestBody.create(mediaType, "{\"flow_run_name\":\"12345转办件\",\"run_name_html\":\"<div contenteditable=\\\"false\\\" class=\\\"title-item\\\">12345转办件</div><div class=\\\"title-item control\\\" data-type=\\\"formData\\\" ng-click=\\\"vm.choiceControl('DATA_2')\\\" data-id=\\\"DATA_2\\\" title=\\\"值来源于-工单编号\\\"></div><div contenteditable=\\\"false\\\" class=\\\"title-item\\\"></div>\",\"flow_id\":\"15\",\"creator\":\"WV00000104\",\"user_name\":\"举报投诉中心\",\"instancy_type\":\"0\",\"form_data\":{\"DATA_31\":\"\",\"DATA_2\":\"\",\"DATA_7\":\"\",\"DATA_3\":\"\",\"DATA_8\":\"\",\"DATA_5\":\"\",\"DATA_10\":\"\",\"DATA_13\":\"\",\"DATA_16\":\"\",\"DATA_27\":\"\",\"DATA_30\":\"\",\"DATA_15\":\"\",\"DATA_17\":\"\",\"DATA_18\":\"\",\"DATA_19\":\"\",\"DATA_20\":\"\",\"DATA_21\":\"\",\"DATA_24\":\"\",\"DATA_25\":\"\"},\"form_structure\":{\"DATA_31\":{\"control_id\":\"DATA_31\",\"control_title\":\"类型\",\"control_type\":\"radio\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":true,\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_31\\\",\\\"title\\\":\\\"\\\\u7c7b\\\\u578b\\\",\\\"data-efb-control\\\":\\\"radio\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u9009\\\\u6846\\\",\\\"data-efb-control-radio\\\":\\\"\\\",\\\"type\\\":\\\"radio\\\",\\\"class\\\":\\\"control-default-radio mceNonEditable\\\",\\\"data-efb-orientation\\\":\\\"h\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"data-efb-options\\\":\\\"\\\\u76f4\\\\u529e\\\\u4ef6,\\\\u627f\\\\u529e\\\\u4ef6\\\",\\\"data-efb-selected-options\\\":\\\"\\\"}\"},\"DATA_2\":{\"control_id\":\"DATA_2\",\"control_title\":\"工单编号\",\"control_type\":\"text\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":true,\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_2\\\",\\\"title\\\":\\\"\\\\u5de5\\\\u5355\\\\u7f16\\\\u53f7\\\",\\\"data-efb-control\\\":\\\"text\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u884c\\\\u6587\\\\u672c\\\\u6846\\\",\\\"data-efb-control-text\\\":\\\"\\\",\\\"type\\\":\\\"text\\\",\\\"class\\\":\\\"control-default-text mceNonEditable\\\",\\\"data-efb-border-type\\\":\\\"all\\\",\\\"data-efb-format\\\":\\\"text\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"data-efb-datetime-calculate\\\":\\\"{\\\\\\\"format\\\\\\\":\\\\\\\"day\\\\\\\",\\\\\\\"formatDecimalPlace\\\\\\\":null}\\\",\\\"style\\\":\\\"height: 27px;\\\",\\\"data-mce-style\\\":\\\"height: 27px;\\\"}\"},\"DATA_7\":{\"control_id\":\"DATA_7\",\"control_title\":\"办结时限\",\"control_type\":\"text\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":true,\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_7\\\",\\\"title\\\":\\\"\\\\u529e\\\\u7ed3\\\\u65f6\\\\u9650\\\",\\\"data-efb-control\\\":\\\"text\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u884c\\\\u6587\\\\u672c\\\\u6846\\\",\\\"data-efb-control-text\\\":\\\"\\\",\\\"type\\\":\\\"text\\\",\\\"class\\\":\\\"control-default-text mceNonEditable\\\",\\\"data-efb-border-type\\\":\\\"all\\\",\\\"data-efb-format\\\":\\\"text\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"data-efb-datetime-calculate\\\":\\\"{\\\\\\\"format\\\\\\\":\\\\\\\"day\\\\\\\",\\\\\\\"formatDecimalPlace\\\\\\\":null}\\\",\\\"style\\\":\\\"height: 27px;\\\",\\\"data-mce-style\\\":\\\"height: 27px;\\\"}\"},\"DATA_3\":{\"control_id\":\"DATA_3\",\"control_title\":\"来电类别\",\"control_type\":\"text\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":true,\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_3\\\",\\\"title\\\":\\\"\\\\u6765\\\\u7535\\\\u7c7b\\\\u522b\\\",\\\"data-efb-control\\\":\\\"text\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u884c\\\\u6587\\\\u672c\\\\u6846\\\",\\\"data-efb-control-text\\\":\\\"\\\",\\\"type\\\":\\\"text\\\",\\\"class\\\":\\\"control-default-text mceNonEditable\\\",\\\"data-efb-border-type\\\":\\\"all\\\",\\\"data-efb-format\\\":\\\"text\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"style\\\":\\\"height: 27px;\\\",\\\"data-efb-datetime-calculate\\\":\\\"{\\\\\\\"format\\\\\\\":\\\\\\\"day\\\\\\\",\\\\\\\"formatDecimalPlace\\\\\\\":null}\\\",\\\"data-mce-style\\\":\\\"height: 27px;\\\"}\"},\"DATA_8\":{\"control_id\":\"DATA_8\",\"control_title\":\"紧急程度\",\"control_type\":\"text\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":true,\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_8\\\",\\\"title\\\":\\\"\\\\u7d27\\\\u6025\\\\u7a0b\\\\u5ea6\\\",\\\"data-efb-control\\\":\\\"text\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u884c\\\\u6587\\\\u672c\\\\u6846\\\",\\\"data-efb-control-text\\\":\\\"\\\",\\\"type\\\":\\\"text\\\",\\\"class\\\":\\\"control-default-text mceNonEditable\\\",\\\"data-efb-border-type\\\":\\\"all\\\",\\\"data-efb-format\\\":\\\"text\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"data-efb-datetime-calculate\\\":\\\"{\\\\\\\"format\\\\\\\":\\\\\\\"day\\\\\\\",\\\\\\\"formatDecimalPlace\\\\\\\":null}\\\",\\\"style\\\":\\\"height: 27px;\\\",\\\"data-mce-style\\\":\\\"height: 27px;\\\"}\"},\"DATA_5\":{\"control_id\":\"DATA_5\",\"control_title\":\"联系人\",\"control_type\":\"text\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":true,\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_5\\\",\\\"title\\\":\\\"\\\\u8054\\\\u7cfb\\\\u4eba\\\",\\\"data-efb-control\\\":\\\"text\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u884c\\\\u6587\\\\u672c\\\\u6846\\\",\\\"data-efb-control-text\\\":\\\"\\\",\\\"type\\\":\\\"text\\\",\\\"class\\\":\\\"control-default-text mceNonEditable\\\",\\\"data-efb-border-type\\\":\\\"all\\\",\\\"data-efb-format\\\":\\\"text\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"data-efb-datetime-calculate\\\":\\\"{\\\\\\\"format\\\\\\\":\\\\\\\"day\\\\\\\",\\\\\\\"formatDecimalPlace\\\\\\\":null}\\\",\\\"style\\\":\\\"height: 27px;\\\",\\\"data-mce-style\\\":\\\"height: 27px;\\\"}\"},\"DATA_10\":{\"control_id\":\"DATA_10\",\"control_title\":\"联系电话\",\"control_type\":\"text\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":true,\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_10\\\",\\\"title\\\":\\\"\\\\u8054\\\\u7cfb\\\\u7535\\\\u8bdd\\\",\\\"data-efb-control\\\":\\\"text\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u884c\\\\u6587\\\\u672c\\\\u6846\\\",\\\"data-efb-control-text\\\":\\\"\\\",\\\"type\\\":\\\"text\\\",\\\"class\\\":\\\"control-default-text mceNonEditable\\\",\\\"data-efb-border-type\\\":\\\"all\\\",\\\"data-efb-format\\\":\\\"text\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"data-efb-datetime-calculate\\\":\\\"{\\\\\\\"format\\\\\\\":\\\\\\\"day\\\\\\\",\\\\\\\"formatDecimalPlace\\\\\\\":null}\\\",\\\"style\\\":\\\"height: 27px;\\\",\\\"data-mce-style\\\":\\\"height: 27px;\\\"}\"},\"DATA_13\":{\"control_id\":\"DATA_13\",\"control_title\":\"问题分类\",\"control_type\":\"text\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":true,\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_13\\\",\\\"title\\\":\\\"\\\\u95ee\\\\u9898\\\\u5206\\\\u7c7b\\\",\\\"data-efb-control\\\":\\\"text\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u884c\\\\u6587\\\\u672c\\\\u6846\\\",\\\"data-efb-control-text\\\":\\\"\\\",\\\"type\\\":\\\"text\\\",\\\"class\\\":\\\"control-default-text mceNonEditable\\\",\\\"data-efb-border-type\\\":\\\"all\\\",\\\"data-efb-format\\\":\\\"text\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"data-efb-datetime-calculate\\\":\\\"{\\\\\\\"format\\\\\\\":\\\\\\\"day\\\\\\\",\\\\\\\"formatDecimalPlace\\\\\\\":null}\\\",\\\"style\\\":\\\"height: 27px; width: 500px;\\\",\\\"data-efb-width\\\":\\\"500px\\\",\\\"data-mce-style\\\":\\\"height: 27px; width: 500px;\\\"}\"},\"DATA_16\":{\"control_id\":\"DATA_16\",\"control_title\":\"问题描述\",\"control_type\":\"textarea\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":true,\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_16\\\",\\\"title\\\":\\\"\\\\u95ee\\\\u9898\\\\u63cf\\\\u8ff0\\\",\\\"data-efb-control\\\":\\\"textarea\\\",\\\"data-efb-control-name\\\":\\\"\\\\u591a\\\\u884c\\\\u6587\\\\u672c\\\\u6846\\\",\\\"data-efb-control-textarea\\\":\\\"\\\",\\\"class\\\":\\\"control-default-textarea mceNonEditable\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"data-efb-width\\\":\\\"500px\\\",\\\"style\\\":\\\"width: 500px;\\\",\\\"data-mce-style\\\":\\\"width: 500px;\\\"}\"},\"DATA_27\":{\"control_id\":\"DATA_27\",\"control_title\":\"问题核实情况选择\",\"control_type\":\"radio\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":true,\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_27\\\",\\\"title\\\":\\\"\\\\u95ee\\\\u9898\\\\u6838\\\\u5b9e\\\\u60c5\\\\u51b5\\\\u9009\\\\u62e9\\\",\\\"data-efb-control\\\":\\\"radio\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u9009\\\\u6846\\\",\\\"data-efb-control-radio\\\":\\\"\\\",\\\"type\\\":\\\"radio\\\",\\\"class\\\":\\\"control-default-radio mceNonEditable\\\",\\\"data-efb-orientation\\\":\\\"h\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"data-efb-options\\\":\\\"\\\\u7ecf\\\\u6838\\\\u5b9e\\\\u5b9e\\\\u9645\\\\u60c5\\\\u51b5\\\\u4e0e12345\\\\u63cf\\\\u8ff0\\\\u4e00\\\\u81f4,\\\\u7ecf\\\\u6838\\\\u5b9e\\\\u5b9e\\\\u9645\\\\u60c5\\\\u51b5\\\\u4e0e12345\\\\u63cf\\\\u8ff0\\\\u4e0d\\\\u4e00\\\\u81f4\\\",\\\"data-efb-selected-options\\\":\\\"\\\"}\"},\"DATA_30\":{\"control_id\":\"DATA_30\",\"control_title\":\"问题核实情况描述\",\"control_type\":\"text\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":true,\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_30\\\",\\\"title\\\":\\\"\\\\u95ee\\\\u9898\\\\u6838\\\\u5b9e\\\\u60c5\\\\u51b5\\\\u63cf\\\\u8ff0\\\",\\\"data-efb-control\\\":\\\"text\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u884c\\\\u6587\\\\u672c\\\\u6846\\\",\\\"data-efb-control-text\\\":\\\"\\\",\\\"type\\\":\\\"text\\\",\\\"class\\\":\\\"control-default-text mceNonEditable\\\",\\\"data-efb-border-type\\\":\\\"all\\\",\\\"data-efb-format\\\":\\\"text\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"data-efb-datetime-calculate\\\":\\\"{\\\\\\\"format\\\\\\\":\\\\\\\"day\\\\\\\",\\\\\\\"formatDecimalPlace\\\\\\\":null}\\\",\\\"style\\\":\\\"height: 100px; width: 500px;\\\",\\\"data-efb-width\\\":\\\"500px\\\",\\\"data-mce-style\\\":\\\"height: 100px; width: 500px;\\\"}\"},\"DATA_15\":{\"control_id\":\"DATA_15\",\"control_title\":\"举报投诉中心处理意见\",\"control_type\":\"text\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":true,\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_15\\\",\\\"title\\\":\\\"\\\\u4e3e\\\\u62a5\\\\u6295\\\\u8bc9\\\\u4e2d\\\\u5fc3\\\\u5904\\\\u7406\\\\u610f\\\\u89c1\\\",\\\"data-efb-control\\\":\\\"text\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u884c\\\\u6587\\\\u672c\\\\u6846\\\",\\\"data-efb-control-text\\\":\\\"\\\",\\\"type\\\":\\\"text\\\",\\\"class\\\":\\\"control-default-text mceNonEditable\\\",\\\"data-efb-border-type\\\":\\\"all\\\",\\\"data-efb-format\\\":\\\"text\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"data-efb-datetime-calculate\\\":\\\"{\\\\\\\"format\\\\\\\":\\\\\\\"day\\\\\\\",\\\\\\\"formatDecimalPlace\\\\\\\":null}\\\",\\\"style\\\":\\\"height: 200px; width: 500px;\\\",\\\"data-efb-width\\\":\\\"500px\\\",\\\"data-mce-style\\\":\\\"height: 27px; width: 500px;\\\"}\"},\"DATA_17\":{\"control_id\":\"DATA_17\",\"control_title\":\"举报投诉中心处理意见时间\",\"control_type\":\"text\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":true,\"empty\":true,\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_17\\\",\\\"title\\\":\\\"\\\\u4e3e\\\\u62a5\\\\u6295\\\\u8bc9\\\\u4e2d\\\\u5fc3\\\\u5904\\\\u7406\\\\u610f\\\\u89c1\\\\u65f6\\\\u95f4\\\",\\\"data-efb-control\\\":\\\"text\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u884c\\\\u6587\\\\u672c\\\\u6846\\\",\\\"data-efb-control-text\\\":\\\"\\\",\\\"type\\\":\\\"text\\\",\\\"class\\\":\\\"control-default-text mceNonEditable\\\",\\\"data-efb-border-type\\\":\\\"all\\\",\\\"data-efb-format\\\":\\\"datetime\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"style\\\":\\\"height: 27px;\\\",\\\"data-efb-default\\\":\\\"\\\",\\\"data-mce-style\\\":\\\"height: 27px;\\\",\\\"data-efb-datetime-calculate\\\":\\\"{\\\\\\\"format\\\\\\\":\\\\\\\"day\\\\\\\",\\\\\\\"formatDecimalPlace\\\\\\\":null}\\\",\\\"data-efb-source\\\":\\\"currentData\\\",\\\"data-efb-source-value\\\":\\\"datetime_dateTime\\\"}\"},\"DATA_18\":{\"control_id\":\"DATA_18\",\"control_title\":\"劳动监察和调解仲裁科审核意见\",\"control_type\":\"text\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":\"\",\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_18\\\",\\\"title\\\":\\\"\\\\u52b3\\\\u52a8\\\\u76d1\\\\u5bdf\\\\u548c\\\\u8c03\\\\u89e3\\\\u4ef2\\\\u88c1\\\\u79d1\\\\u5ba1\\\\u6838\\\\u610f\\\\u89c1\\\",\\\"data-efb-control\\\":\\\"text\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u884c\\\\u6587\\\\u672c\\\\u6846\\\",\\\"data-efb-control-text\\\":\\\"\\\",\\\"type\\\":\\\"text\\\",\\\"class\\\":\\\"control-default-text mceNonEditable\\\",\\\"data-efb-border-type\\\":\\\"all\\\",\\\"data-efb-format\\\":\\\"text\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"data-efb-datetime-calculate\\\":\\\"{\\\\\\\"format\\\\\\\":\\\\\\\"day\\\\\\\",\\\\\\\"formatDecimalPlace\\\\\\\":null}\\\",\\\"style\\\":\\\"height: 27px; width: 500px;\\\",\\\"data-efb-width\\\":\\\"500px\\\",\\\"data-mce-style\\\":\\\"height: 27px; width: 500px;\\\"}\"},\"DATA_19\":{\"control_id\":\"DATA_19\",\"control_title\":\"劳动监察和调解仲裁科审核意见时间\",\"control_type\":\"text\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":\"\",\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_19\\\",\\\"title\\\":\\\"\\\\u52b3\\\\u52a8\\\\u76d1\\\\u5bdf\\\\u548c\\\\u8c03\\\\u89e3\\\\u4ef2\\\\u88c1\\\\u79d1\\\\u5ba1\\\\u6838\\\\u610f\\\\u89c1\\\\u65f6\\\\u95f4\\\",\\\"data-efb-control\\\":\\\"text\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u884c\\\\u6587\\\\u672c\\\\u6846\\\",\\\"data-efb-control-text\\\":\\\"\\\",\\\"type\\\":\\\"text\\\",\\\"class\\\":\\\"control-default-text mceNonEditable\\\",\\\"data-efb-border-type\\\":\\\"all\\\",\\\"data-efb-format\\\":\\\"datetime\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"style\\\":\\\"height: 27px;\\\",\\\"data-efb-default\\\":\\\"\\\",\\\"data-mce-style\\\":\\\"height: 27px;\\\",\\\"data-efb-datetime-calculate\\\":\\\"{\\\\\\\"format\\\\\\\":\\\\\\\"day\\\\\\\",\\\\\\\"formatDecimalPlace\\\\\\\":null}\\\",\\\"data-efb-source\\\":\\\"currentData\\\",\\\"data-efb-source-value\\\":\\\"datetime_dateTime\\\"}\"},\"DATA_20\":{\"control_id\":\"DATA_20\",\"control_title\":\"分管领导意见\",\"control_type\":\"text\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":\"\",\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_20\\\",\\\"title\\\":\\\"\\\\u5206\\\\u7ba1\\\\u9886\\\\u5bfc\\\\u610f\\\\u89c1\\\",\\\"data-efb-control\\\":\\\"text\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u884c\\\\u6587\\\\u672c\\\\u6846\\\",\\\"data-efb-control-text\\\":\\\"\\\",\\\"type\\\":\\\"text\\\",\\\"class\\\":\\\"control-default-text mceNonEditable\\\",\\\"data-efb-border-type\\\":\\\"all\\\",\\\"data-efb-format\\\":\\\"text\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"data-efb-datetime-calculate\\\":\\\"{\\\\\\\"format\\\\\\\":\\\\\\\"day\\\\\\\",\\\\\\\"formatDecimalPlace\\\\\\\":null}\\\",\\\"style\\\":\\\"height: 27px; width: 500px;\\\",\\\"data-efb-width\\\":\\\"500px\\\",\\\"data-mce-style\\\":\\\"height: 27px; width: 500px;\\\"}\"},\"DATA_21\":{\"control_id\":\"DATA_21\",\"control_title\":\"分管领导意见时间\",\"control_type\":\"text\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":\"\",\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_21\\\",\\\"title\\\":\\\"\\\\u5206\\\\u7ba1\\\\u9886\\\\u5bfc\\\\u610f\\\\u89c1\\\\u65f6\\\\u95f4\\\",\\\"data-efb-control\\\":\\\"text\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u884c\\\\u6587\\\\u672c\\\\u6846\\\",\\\"data-efb-control-text\\\":\\\"\\\",\\\"type\\\":\\\"text\\\",\\\"class\\\":\\\"control-default-text mceNonEditable\\\",\\\"data-efb-border-type\\\":\\\"all\\\",\\\"data-efb-format\\\":\\\"datetime\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"style\\\":\\\"height: 27px;\\\",\\\"data-efb-default\\\":\\\"\\\",\\\"data-mce-style\\\":\\\"height: 27px;\\\",\\\"data-efb-datetime-calculate\\\":\\\"{\\\\\\\"format\\\\\\\":\\\\\\\"day\\\\\\\",\\\\\\\"formatDecimalPlace\\\\\\\":null}\\\",\\\"data-efb-source\\\":\\\"currentData\\\",\\\"data-efb-source-value\\\":\\\"datetime_dateTime\\\"}\"},\"DATA_24\":{\"control_id\":\"DATA_24\",\"control_title\":\"承办科室处理结果\",\"control_type\":\"text\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":\"\",\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_24\\\",\\\"title\\\":\\\"\\\\u627f\\\\u529e\\\\u79d1\\\\u5ba4\\\\u5904\\\\u7406\\\\u7ed3\\\\u679c\\\",\\\"data-efb-control\\\":\\\"text\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u884c\\\\u6587\\\\u672c\\\\u6846\\\",\\\"data-efb-control-text\\\":\\\"\\\",\\\"type\\\":\\\"text\\\",\\\"class\\\":\\\"control-default-text mceNonEditable\\\",\\\"data-efb-border-type\\\":\\\"all\\\",\\\"data-efb-format\\\":\\\"text\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"data-efb-datetime-calculate\\\":\\\"{\\\\\\\"format\\\\\\\":\\\\\\\"day\\\\\\\",\\\\\\\"formatDecimalPlace\\\\\\\":null}\\\",\\\"style\\\":\\\"height: 200px; width: 500px;\\\",\\\"data-efb-width\\\":\\\"500px\\\",\\\"data-mce-style\\\":\\\"height: 27px; width: 500px;\\\"}\"},\"DATA_25\":{\"control_id\":\"DATA_25\",\"control_title\":\"承办科室处理结果时间\",\"control_type\":\"text\",\"control_parent_id\":\"\",\"control_purview\":{\"edit\":\"\",\"empty\":\"\",\"always\":\"\",\"countersignVisible\":{\"flag\":\"\",\"nodeId\":\"66\"}},\"control_attribute\":\"{\\\"id\\\":\\\"DATA_25\\\",\\\"title\\\":\\\"\\\\u627f\\\\u529e\\\\u79d1\\\\u5ba4\\\\u5904\\\\u7406\\\\u7ed3\\\\u679c\\\\u65f6\\\\u95f4\\\",\\\"data-efb-control\\\":\\\"text\\\",\\\"data-efb-control-name\\\":\\\"\\\\u5355\\\\u884c\\\\u6587\\\\u672c\\\\u6846\\\",\\\"data-efb-control-text\\\":\\\"\\\",\\\"type\\\":\\\"text\\\",\\\"class\\\":\\\"control-default-text mceNonEditable\\\",\\\"data-efb-border-type\\\":\\\"all\\\",\\\"data-efb-format\\\":\\\"datetime\\\",\\\"contenteditable\\\":\\\"false\\\",\\\"style\\\":\\\"height: 27px;\\\",\\\"data-efb-default\\\":\\\"\\\",\\\"data-mce-style\\\":\\\"height: 27px;\\\",\\\"data-efb-datetime-calculate\\\":\\\"{\\\\\\\"format\\\\\\\":\\\\\\\"day\\\\\\\",\\\\\\\"formatDecimalPlace\\\\\\\":null}\\\",\\\"data-efb-source\\\":\\\"currentData\\\",\\\"data-efb-source-value\\\":\\\"datetime_dateTime\\\"}\"}}}");
        Request request = new Request.Builder()
                .url("http://120.221.150.148:8010/eoffice10/server/public/api/flow/new-page/flow-save")
                .method("POST", body)
                .addHeader("Accept", "application/json, text/plain, */*")
                .addHeader("Accept-Encoding", "gzip, deflate")
                .addHeader("Accept-Language", "zh-CN,zh;q=0.9")
                .addHeader("Authorization", "Bearer " + token)
                .addHeader("Connection", "keep-alive")
                .addHeader("Content-Type", "application/json;charset=UTF-8")
//                .addHeader("Cookie", "io=brfo2wKmVloOIqBUAEDf")
                .addHeader("Host", "120.221.150.148:8010")
                .addHeader("Origin", "http://120.221.150.148:8010")
                .addHeader("Referer", "http://120.221.150.148:8010/eoffice10/client/app/web/")
                .addHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.25 Safari/537.36 Core/1.70.3756.400 QQBrowser/10.5.4039.400")
                .addHeader("Content-Length", "16081")
                .build();
        Response response = client.newCall(request).execute();
        String run_id = JSONObject.parseObject(JSONObject.parseObject(response.body().string()).getString("data")).getString("run_id");
        System.out.println("获取run_id成功！");
        return run_id;
    }

    /*
    写数据：通过修改流程写入数据
    */
    public static void inputOA(String token, String content) throws IOException {
        OkHttpClient client = new OkHttpClient().newBuilder()
                .build();
        MediaType mediaType = MediaType.parse("application/json;charset=UTF-8");
        RequestBody body = RequestBody.create(mediaType, content);
        Request request = new Request.Builder()
                .url("http://120.221.150.148:8010/eoffice10/server/public/api/flow/run/save-flow-run-info")
                .method("POST", body)
                .addHeader("Accept", "application/json, text/plain, */*")
                .addHeader("Accept-Encoding", "gzip, deflate")
                .addHeader("Accept-Language", "zh-CN,zh;q=0.9")
                .addHeader("Authorization", "Bearer " + token)
                .addHeader("Connection", "keep-alive")
                .addHeader("Content-Type", "application/json;charset=UTF-8")
//                .addHeader("Cookie", "io=iWbkGUXbzfy4s0_XAD_K")
                .addHeader("Host", "120.221.150.148:8010")
                .addHeader("Content-Length", content.getBytes("UTF-8").length + "")
                .build();
        Response response = client.newCall(request).execute();
        String run_id = JSONObject.parseObject(JSONObject.parseObject(response.body().string()).getString("data")).getString("run_id");
        System.out.println("写入成功！run_id=" + run_id);
    }

    /*
    type:直办件、转办件、退办件
     */
    public static String getContent(String run_id, String order_code, String phone_type, String link_person, String end_date, String urgency_degree, String link_phone,
                                    String problem_classification, String problem_description, String transfer_process, String suggestion, String type) {
        Date dNow = new Date();
        SimpleDateFormat ft = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
        String temp = "{\"run_id\":\"${run_id}\",\"run_name\":\"12345转办件${工单编号}\",\"run_name_html\":\"<div contenteditable=\\\"false\\\" class=\\\"title-item\\\">12345转办件</div><div class=\\\"title-item control\\\" data-type=\\\"formData\\\" ng-click=\\\"vm.choiceControl('DATA_2')\\\" data-id=\\\"DATA_2\\\" title=\\\"值来源于-工单编号\\\" ng-bind=\\\"vm.praseData('DATA_2')\\\">${工单编号}</div><div contenteditable=\\\"false\\\" class=\\\"title-item\\\"></div>\",\"instancy_type\":\"${紧急编号}\",\"form_data\":{\"DATA_31\":\"${办件类型}\",\"DATA_2\":\"${工单编号}\",\"DATA_7\":\"${办结时限}\",\"DATA_3\":\"${来电类别}\",\"DATA_8\":\"${紧急程度}\",\"DATA_5\":\"${联系人}\",\"DATA_10\":\"${联系电话}\",\"DATA_13\":\"${问题分类}\",\"DATA_16\":\"${问题描述}\",\"DATA_27\":\"经核实实际情况与12345描述一致\",\"DATA_30\":\"${问题核实情况}\",\"DATA_15\":\"${处理意见}\",\"DATA_17\":\"${处理意见时间}\",\"DATA_18\":\"\",\"DATA_19\":\"\",\"DATA_20\":\"\",\"DATA_21\":\"\",\"DATA_24\":\"\",\"DATA_25\":\"\"},\"flow_process\":66,\"process_id\":1}";
        String str = temp.replace("${run_id}", run_id)
                .replace("${工单编号}", order_code)
                .replace("${来电类别}", phone_type)
                .replace("${联系人}", link_person)
                .replace("${办结时限}", end_date)
                .replace("${紧急程度}", urgency_degree)
                .replace("${联系电话}", link_phone)
                .replace("${问题分类}", problem_classification)
                .replace("${问题描述}", problem_description)
                .replace("${问题核实情况}", transfer_process)
                .replace("${处理意见}", suggestion)
                .replace("${办件类型}", type)
                .replace("${处理意见时间}", ft.format(dNow));
        if (type.equals("退办件")) {
            str.replace("${紧急编号}", "0").replace("${办件类型}", "直办件");
        } else if (type.equals("转办件")) {
            str.replace("${办件类型}", "承办件");
        } else {
            str.replace("${办件类型}", type);
        }
        if (urgency_degree.equals("一般")) {
            str.replace("${紧急编号}", "0");
        } else {
            str.replace("${紧急编号}", "2");
        }
        System.out.println(run_id + "-" + order_code + "-" + link_person + "-" + type + "-" + end_date);
        return str;
    }

    /*
    读取单个文档内容并写入
    */
    public static void readDocx(String filePath, String token) {
        try {
            FileInputStream in = new FileInputStream(filePath);
            XWPFDocument xwpf = new XWPFDocument(in);
            Iterator<XWPFTable> it = xwpf.getTablesIterator();
            XWPFTable table = it.next();
            List<XWPFTableRow> rows = table.getRows();
            String end_date = rows.get(0).getTableCells().get(3).getText();
            String order_code = rows.get(1).getTableCells().get(1).getText();
            String urgency_degree = rows.get(1).getTableCells().get(3).getText();
            String phone_type = rows.get(2).getTableCells().get(1).getText();
            String link_person = rows.get(4).getTableCells().get(1).getText();
            String link_phone = rows.get(4).getTableCells().get(3).getText();
            String problem_classification = rows.get(7).getTableCells().get(1).getText();
            String problem_description = rows.get(8).getTableCells().get(1).getText();
            String transfer_process = rows.get(9).getTableCells().get(1).getText();
            String type = rows.get(12).getTableCells().get(1).getText();
            String department = rows.get(12).getTableCells().get(3).getText();
            String suggestion = "建议转" + department + "进行答复。";
            String run_id = getRun_id(token);
            String content = getContent(run_id, order_code, phone_type, link_person, end_date, urgency_degree, link_phone, problem_classification, problem_description, transfer_process, suggestion, type);
            inputOA(token, content);
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("文档样式有误！");
        }
    }


    public static void main(String[] args) {
        try {
            System.out.println("所有与文件必须在D盘名为“上传OA”的文件夹下面！");
            System.out.println("文件后缀必须是docx！");
            System.out.println("文件并非12345默认下载的样式，模板有调整！");
            String token = getToken();
            File[] files = new File("D:\\上传OA").listFiles();
            if (files != null) {
                for (File file : files) {
                    String fileName = file.getName();
                    if (fileName.endsWith("docx")) {
                        readDocx(file.getAbsolutePath(), token);
                    }
                }

            }
            System.out.println("运行完毕，请登录OA提交流程！");
            Scanner input = new Scanner(System.in);
            String val = null;       // 记录输入度的字符串
            val = input.next();       // 等待输入值

        } catch (Exception e) {
            System.out.println("错误！请确认网络情况！");
            e.printStackTrace();
        }
    }
}
