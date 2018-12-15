import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import okhttp3.*;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {

    private  String cookies =
            "safedog-flow-item=; ASP.NET_SessionId=s5uvidfgpanjo0tdoqsrrtzv; YangHua=E84AF1365B39D0A3C8B623043B12913BD0880527AB836438294EAB8BA72CBFF30B05762A2A9B4217E808A2CF21BC6EB4E10A4AF2EE997D8EE405BD11409662DBA8EE225414E1B4ADC5BE310726C726BF3B738EFB47CDEF78106453D3818ABF4C6E82F247D790228D92B72DC9A414555B0949927E1076CCE96F13411BF8D0C39F937F2353763623957D19DADF11AD3616AC43F2B2BE008292B8AE5E104600D33E49F7BD48";

    private   int CollegeNoNum = 0;
    private   int GradeNum = 0;
    private   int QNum = 17;
      List<DataItem> dataItemList = new ArrayList<>();


    private   String[] GradeAll = {"2015","2016","2017","2018"};
    private   String[] CollegeNoAll = {
            "01",
            "02",
            "04",
            "05",
            "07",
            "08",
            "09",
            "10",
            "11",
            "12",
            "13",
            "14",
            "15",
            "16",
            "17",
            "19",
            "20",
            "21",
            "24",
            "26",
            "28",
            "29",
            "30",
            "22",
            "06",
            "03",
            "32",
            "18",
            "23",
            "27",
            "31"};
    private   String[] CollegeAll = {
            "信息科学与工程学院",
            "物理科学与技术学院",
            "资源环境学院",
            "化学化工学院",
            "管理学院",
            "生命科学学院",
            "土木工程与力学学院",
            "外国语学院",
            "大气科学学院",
            "经济学院",
            "新闻与传播学院",
            "文学院",
            "历史文化学院",
            "核科学与技术学院",
            "数学与统计学院",
            "艺术学院",
            "马克思主义学院",
            "法学院",
            "哲学社会学院",
            "教育学院",
            "地质科学与矿产资源学院",
            "萃英学院",
            "体育部",
            "药学院",
            "基础医学院",
            "第一临床医学院",
            "政治与国际关系学院",
            "第二临床医学院",
            "公共卫生学院",
            "口腔医学院",
            "护理学院"};


    private   String cookies1;
    private   String cookies2;
    private   String cookies3;
    public  static void main(String[] args) {
        System.out.println("Hello World!");
        Main aa = new Main();

        aa.GetXGWDCWJ0();
//        File input = new File("D:/a.html");
//        GetInf(input);

    }



    /**
     * 学工网调查问卷
     */

    private   void GetXGWDCWJ0(){

        OkHttpClient okHttpClient = new OkHttpClient
                .Builder()
                .followRedirects(false)//！！！！！！！重定向
                .followSslRedirects(false)
                .connectTimeout(100, TimeUnit.SECONDS) //设置连接超时
                .readTimeout(100, TimeUnit.SECONDS) //设置读超时
                .writeTimeout(100, TimeUnit.SECONDS) //设置写超时
                .retryOnConnectionFailure(true) //是否自动重连
                .build(); //构建OkHttpClient对象

        Request request = new Request.Builder()
                .url("http://xgb.lzu.edu.cn/SystemForm/OnLine/VoteReport.aspx?Id=32&Form=OnLineTopic.aspx")
                .addHeader("Cookie", cookies)
                .build();

        Call call = okHttpClient.newCall(request);
        call.enqueue(new Callback() {
            @Override
            public void onFailure( Call call,  IOException e) {
                System.out.println("info_call0fail\n"+e.toString());
            }

            @Override
            public void onResponse( Call call,  Response response) throws IOException {

                if (response.code() == 200) {
                    String results = response.body().string();
//                    System.out.println(results);
                    GetInf0(results);
                }
            }
        });
    }

    private   void GetInf0(String code){
        Document doc = Jsoup.parse(code, "UTF-8");

        cookies2 = doc.select("input").get(0).attr("value");
        cookies1 = doc.select("input").get(2).attr("value");
        cookies3 = doc.select("input").get(1).attr("value");

//        cookies2 = doc.getElementsByClass("aspNetHidden").get(0).select("input").get(2).attr("value");
//        cookies1 = doc.getElementsByClass("aspNetHidden").get(1).select("input").get(1).attr("value");
//        cookies3 = doc.getElementsByClass("aspNetHidden").get(1).select("input").get(0).attr("value");
//        System.out.println(cookies1);
//        System.out.println(cookies2);
//        System.out.println(cookies3);

            CollegeNoNum = 0;
                GradeNum = 0;

        PostXGWDCWJ(CollegeNoNum,GradeNum,"32");

//        for (int i = 0; i < 5; i++){
//            for (int j = 0; j < GradeAll.length; j++){
//
//                PostXGWDCWJ(i,j,"32");
//
//            }
//        }


//        for (int i = 0; i < CollegeNoAll.length; i++){
//            for (int j = 0; j < GradeAll.length; j++){
//
//                PostXGWDCWJ(i,j,"32");
//            }
//        }




    }

    private   void PostXGWDCWJ(int ii,int jj,String id){


        String CollegeNo = CollegeNoAll[ii];
        String Grade = GradeAll[jj];
        System.out.println(CollegeAll[ii]+GradeAll[jj]);

        String url = "http://xgb.lzu.edu.cn/SystemForm/OnLine/VoteReport.aspx?Id="+id+"&Form=OnLineTopic.aspx";


        try {
            Thread.sleep(8000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }


        OkHttpClient okHttpClient = new OkHttpClient
                .Builder()
                .followRedirects(false)//！！！！！！！重定向
                .followSslRedirects(false)
                .connectTimeout(100, TimeUnit.SECONDS) //设置连接超时
                .readTimeout(100, TimeUnit.SECONDS) //设置读超时
                .writeTimeout(100, TimeUnit.SECONDS) //设置写超时
                .retryOnConnectionFailure(true) //是否自动重连
                .build(); //构建OkHttpClient对象

        RequestBody body = new FormBody.Builder()
                .add("__EVENTTARGET", "")
                .add("__EVENTARGUMENT", "")
                .add("__LASTFOCUS", "")
//                .add("__VIEWSTATE", a)
//                .add("__VIEWSTATEGENERATOR", a2)
//                .add("__EVENTVALIDATION", a3)
                .add("__VIEWSTATE", cookies2)
                .add("__VIEWSTATEGENERATOR", cookies3)
                .add("__EVENTVALIDATION", cookies1)
                .add("SpeType", "-1")
                .add("CollegeNo", CollegeNo)
                .add("Grade", Grade)
                .add("SpecialtyNo", "-1")
                .add("ClassNo", "-1")
                .add("National", "-1")
                .add("Polity", "-1")
                .add("Sex", "")
                .add("Button2", "%D6%B4%D0%D0%CD%B3%BC%C6")
                .build();

        Request request = new Request
                .Builder()
                .url(url)
                .addHeader("Cookie", cookies)
                .addHeader("Host", "xgb.lzu.edu.cn")
                .addHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:63.0) Gecko/20100101 Firefox/63.0")
                .addHeader("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8")
                .addHeader("Accept-Encoding", "gzip, deflate")
                .addHeader("Content-Type", "application/x-www-form-urlencoded")
                .addHeader("Content-Length", "27449")
                .addHeader("Pragma", "no-cache")
                .addHeader("Cache-Control", "no-cache")
                .post(body)
                .build();

//        System.out.println(cookies1);
//        System.out.println(cookies2);
//        System.out.println(cookies3);
//        System.out.println(CollegeNo+CollegeAll[CollegeNoNum]+Grade);
//        System.out.println(url);
//        System.out.println(((FormBody) body).size());

        Call call = okHttpClient.newCall(request);

        call.enqueue(new Callback() {
            @Override
            public void onFailure( Call call,  IOException e) {
                System.out.println("info_call2fail\n"+e.toString());
                System.out.println(CollegeAll[ii]+GradeAll[jj]+"========数据请求错误：");


                DataItem dataItem = new DataItem();
                dataItem.setCollege(CollegeAll[ii]);
                dataItem.setGradle(GradeAll[jj]);
                dataItem.setPeople("0");
                List<String> ques = new ArrayList<>();
                for (int m = 0; m < 68; m++){
                    ques.add("0");
                }
                dataItem.setQue(ques);
                dataItemList.add(dataItem);

                Next();

            }

            @Override
            public void onResponse( Call call,  Response response) throws IOException {

                String results = response.body().string();
//                System.out.println(results);
//                System.out.println(response.code());

                if (response.code() == 200) {
                    GetInf(ii,jj,results);
                }
            }

//
        });

    }

    private   void GetInf(int ii,int jj,String code){

        Document doc = Jsoup.parse(code, "UTF-8");

        List<String> ques = new ArrayList<>();

        String ALLPeopleNum = "0";
        try {

            ALLPeopleNum = doc.getElementById("PersonNum").text();
            Elements Content = doc.getElementById("Content").select("span > table > tbody > tr");
            int trNum = Content.size();

            System.out.println(CollegeAll[ii]+GradeAll[jj]+"========回答人数："+ALLPeopleNum+",问题数：" + trNum/2);

            for (int i = 1; i < trNum; i=i+2) {

                Elements Q = Content.get(i).select("tbody > tr");

                String QTitle = Content.get(i - 1).text();
//                System.out.println("\n"+QTitle);

                for (int j = 0; j < Q.size(); j++) {

                    try {
                        String QItemPeopleNum = getMatcher("\\[(\\d*)人\\]", Q.select("font").get(j).text());
                        String QItem = Q.get(j).text();

                        ques.add(QItemPeopleNum);

                    } catch (Exception e) {
                        e.printStackTrace();

                    }
                }
                if (Q.size()==3){
                    ques.add("0");
                }

            }
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println(CollegeAll[ii]+GradeAll[jj]+"========解析错误：");

        }

        DataItem dataItem = new DataItem();
        dataItem.setCollege(CollegeAll[ii]);
        dataItem.setGradle(GradeAll[jj]);
        dataItem.setPeople(ALLPeopleNum);
        dataItem.setQue(ques);
        dataItemList.add(dataItem);

        Next();

    }

    private  void Next(){

        GradeNum = GradeNum + 1;
        if (GradeNum >= 4){
            CollegeNoNum = CollegeNoNum + 1;
            GradeNum = 0;
            if (CollegeNoNum == CollegeNoAll.length){
                try {
                    CreateExcel(dataItemList);
                    System.out.println(dataItemList.get(0).getCollege());
                } catch (IOException | WriteException e) {
                    e.printStackTrace();
                }
            }else {
                PostXGWDCWJ(CollegeNoNum,GradeNum,"32");

            }
        }else {
            PostXGWDCWJ(CollegeNoNum,GradeNum,"32");

        }

    }

    private  String getMatcher(String regex, String source) {
        String result = "";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(source);
        while (matcher.find()) {
            result = matcher.group(1);
        }
        return result;
    }

    private  void CreateExcel(List<DataItem> dataItemList1)
            throws IOException, WriteException {

        //1:创建excel文件

        String path = "D:/兰州大学.xls";

        File file = new File(path);
        file.createNewFile();
        WritableWorkbook workbook = Workbook.createWorkbook(file);

        //3:创建sheet,设置第二三四..个sheet，依次类推即可
        WritableSheet sheet = workbook.createSheet("辅导员投票", 0);
        //4：设置titles
        //5:单元格
        Label label;
        //6:给第一列设置列名
        List<String> Q_titles= new ArrayList<>();

        Q_titles.add("学院");
        Q_titles.add("年级");
        Q_titles.add("答题人数");

        String ABCD[] = {"A", "B", "C", "D"};

        for(int i=0; i < QNum; i++) {
            for (int j = 0; j < 4; j++){
                Q_titles.add((i+1)+ABCD[j]);
            }
        }

        for(int i=0; i<Q_titles.size(); i++){
            //x,y,第一行的列名
            label=new Label(i, 0, Q_titles.get(i));

            //7：添加单元格
            sheet.addCell(label);
        }

        for(int i = 0; i < dataItemList1.size(); i++){

            //x,y,第一行的列名
            label = new Label(0, i+1, dataItemList1.get(i).getCollege());
            sheet.addCell(label);

            label = new Label(1, i+1, dataItemList1.get(i).getGradle());
            sheet.addCell(label);

            label = new Label(2, i+1, dataItemList1.get(i).getPeople());
            sheet.addCell(label);

            List<String> stringList = dataItemList1.get(i).getQue();
            for (int j = 0; j < stringList.size(); j++){

                label = new Label(j+3, i+1, stringList.get(j));

                sheet.addCell(label);

            }

            //7：添加单元格
        }

        //写入数据，一定记得写入数据，不然你都开始怀疑世界了，excel里面啥都没有
        workbook.write();
        //最后一步，关闭工作簿

        workbook.close();
    }




}
