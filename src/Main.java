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
    static String cookies =
            "safedog-flow-item=; ASP.NET_SessionId=s5uvidfgpanjo0tdoqsrrtzv; YangHua=E84AF1365B39D0A3C8B623043B12913BD0880527AB836438294EAB8BA72CBFF30B05762A2A9B4217E808A2CF21BC6EB4E10A4AF2EE997D8EE405BD11409662DBA8EE225414E1B4ADC5BE310726C726BF3B738EFB47CDEF78106453D3818ABF4C6E82F247D790228D92B72DC9A414555B0949927E1076CCE96F13411BF8D0C39F937F2353763623957D19DADF11AD3616AC43F2B2BE008292B8AE5E104600D33E49F7BD48";
    int Max_Grade = 2018;
    int Min_Grade = 2015;
    static List<String> QNum = new ArrayList<>();
    private static String[] Q_titles={"题目标号", "1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17"};

    private static String[] GradeAll = {"2015","2016","2017","2018"};
    private static String[] CollegeNoAll = {
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
    private static String[] CollegeAll = {
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


    private static String cookies1;
    private static String cookies2;
    private static String cookies3;
    public static void main(String[] args) {
        System.out.println("Hello World!");

        GetXGWDCWJ0();
//        File input = new File("D:/a.html");
//        GetInf(input);

    }



    /**
     * 学工网调查问卷
     */

    private static void GetXGWDCWJ0(){

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

    private static void GetInf0(String code){
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

//            CollegeNoNum = 0;
//                GradeNum = 0;
//
//        PostXGWDCWJ(CollegeNoAll[0],GradeAll[0],"32");

        for (int i = 0; i < CollegeNoAll.length; i++){
            for (int j = 0; j < GradeAll.length; j++){

                try
                {
                    Thread.sleep(8000);
                }
                catch (InterruptedException e)
                {
                    e.printStackTrace();
                }

                PostXGWDCWJ(i,j,"32");
            }
        }


    }

    private static void PostXGWDCWJ(int ii,int jj,String id){
        String CollegeNo = CollegeNoAll[ii];
        String Grade = GradeAll[jj];
        System.out.println(CollegeAll[ii]+GradeAll[jj]);

        String a = "g%2BIaSs1OU28c23z4TBX2cLND9IrcIGh%2BvY91cYQkHQEdJ1ZcK2xXlCElH6EVQCeeqS6A1xCq56XHMCLadCGEJhd0hqYT%2FtCnAf%2B5LpR8tbXoDrj%2FiVy9gM0exDC93euGqilzx5cVVzQP1rZ5HhkfCjsXWtspTFDoTAEyqzECYRlIvD40htlOi5F2SBDUjT4GlnxKoXbNtXRWDDRQTdorbkiMz3Rnh0Zt38IpORaiyruW48S%2Bn23GM0anWLHIStUW41eWnxA4cwL2LbGegWjzxdSkc6G4oXMbKPBb0KnPcy50kmSdX%2Fe4KQd%2FwHXXD2NSy4dKq4cY%2Fqsol0KysT%2FXgI%2F7zSiZLL5rNHIahuOM%2FsbmtKaFuINpf5ULVM9qB3%2FZNquxqyWcPwkiolU9lC2WKhSsnr9LhamvtBJzrLNDIXX4VsT2rtZwvoe0a%2BOfkBktexIuYSmNec3ONYL2R%2Fjdbrf4YuFY3qXmKv2AokdpuHHO0FkWX1oxppXz0AE%2BIPo7JsQk0qsMGdvvgt%2Fn9DnAvYfRpndqgVsd4dd9CFt9phAZ4szGHoikT2fZW%2BYArUsdGEmC1jPKwCUFqaYV3qZ9DZiMkZvy45esCnnu4zDyMi4eRT22DcTWCLOgwiZ82kD1oKkKV9t8rQcjfkMAzHs%2Bnogi2hU8IzqjHQ5hpnZmzbtR6KeifBf%2FsahBHYwZc1bMfmpS6VvUqWNZmo5e5Deo2PFuwz2tnYb08Fp90uFlXJ0RHGTHe8qL3IVKXB%2B3rz8JqKggO9JtOx97RARefEKZm%2FpR7mTHaXrLfceEauqZ825DoTbs0Msv3u4ZnGdLzkehJMsy3BBTQlT%2B3sCH3LG5lMXUoE1TpazbCrRtepxzNPPKbJ2cKW7ZgSbyNe4TzQbc41iamkI%2FtZ3fHwHKR4GH3nNEChi4pb%2FgWdz7XYYC5JjbA14UTxsVcJCDSzdZjlzffwB%2F6oxXD9brVi2VkU1RdcRTbbc%2Fc6Aafk050FXFBKJIwYhXo2RwJv%2B%2FJyOpNKkaTJpp61BeYLhCFdoeGIUSova6SY80dq8ZQTqDAAxm01c%2BOmmp2bE2LGHep%2B%2B06RIqdvitv87k%2FkM0QlAD0sHWZeDwejBcrr2ELLGf4hBoOLXHQ8m%2F29vJC97zmXJZh5mNkjgD4lg3mM7FyqctCqCAz5l72dgguuQ2q1Ovlp8HlApj2Vpr10lKiUKjsV74Rt8AHeGjjmYdqeWgMLT8ZJXGWyhik6pmUiViARQQY9FbVisgGMW1PTgYmur2%2BC8sjckRJSrAaB0PY6JOcBzl7r9XcCf6S9qqAuSGSc0gLIVoE8lFXps1zs82FhOXvvIpEkjIGGRGc6hEoKUj0%2Fy9tKzn5DQ%2Fu7R1%2BpobTXr2eK2BcHNahq9YT%2BTXG4GTSh6TjIOFgvkyAQQUsI%2Bh27urTMotmuw7ZbZ6wTFk4uxC6CKguadM80wx402hYfdVkdgy0tIEITJz3tJ2zvrKL%2Bx1rRXglJgWgVkHkC2GJxIJq3OqvCYcIWyKNlN4IYpUp6AAcaw6jyIFSxS%2F1rpPTHzC%2BMCT6SPUFaFHv6xPEQZzINRfgYSe1PpGOgKcjrduSThHwpSDdImAq4E38fCwd%2FvnlhS07cEbFEV2Vfojk8IQn9jxwAe87cHBIwwEpdnKHUsltYNLr25eyouolTv%2BOkna3z8HzS9ixiQ73umP%2F9k3HNuNz8NRprHuPrnmA3qCmenqc9VYmRtCrtDjtexCTCu%2BQHTlcayZGlNNaMLFtd74MecRkFTano9QQCQ9WWaNe9EAQUlkiPsPR07BZs8%2FYXQd3a072lC%2FpNUKifA6Nhc%2B7iSQCQ1a08X%2FjghuI06DKyilrDPbBziT8as%2F9%2FawWhz2ZBFuPahsLJ3EsG5DYmwAJZjDYUzgWQmV72rsfpUtUH961YM39DzRom2UFbG1nk4dV4t4bdHjRkIpiAeZT98WB64ksBME9dleVudUYvvWVJ7MpwpD7s9CazL46V%2BYSP7mqKSieJW5JKpmV2PRLHJlewf6dDX13hGf%2B80Ybt3OaPk9mhjbSDeoo39PKz6z5z8SvCyYTdmJjKZOUMp8zAbxoyh1bar5TNU7knwqO3B36DQ3RhjAnJiIV5Od%2BzNF9gNt7DnkLXx2BORxXQgtDj4Fztk%2ByWrP1Uv6W%2BhwxLxiDZnOPQ8reHg%2F0kE65xLeQZcdOTlfqNNwmSyZ61MiOb3vLQ3zUCe8K%2B3hxsuIbK%2B00T5Nf4Q0I0fg3ObnhDBZBfZbleZWcJs05ZuLcj2HqMxpkw90%2BV5mkgXU%2FfoOKopO9nqkDaXY4dHjbMj%2FqWwWGLDbH6AC74gZK%2F9lczDQ0aTwBLlL6vv6gzof21ppT0n0CRf8lE4%2F9FtPbfUQ0t0V5E2k%2BykGBGB1O4nTAgYKllFlVKNh160uOF1Lah0ZdH120DrMD8e2dldLU%2B%2FktBQHbM5oF1QbAQbGLRsf3NK5B0P5geoUyBZC2fU4%2BJriR4DOYQG2q9A8Qukvp6Tl93JqxqBJMde9IsfHt5eMPraoS0C3JU7aMEdicVkO2SnKb6EsAtBYvHIUxOjxuMUoDC6n0GFI%2B1TRN99%2Bw2KoCLY58Hh1cwalX4rxbGz%2FOom6uS%2FnsDEtZvOcL8Z5dUdlKOHgVpCa9WoBD862vi4pcJoNBgkJLmoRK%2BWBTUc3NPOJww5xwHEa9G4nlObiArToVl0vWS%2Ftud4bzv2ASwtNnMP0HFZ5aA6FQ1orDVYNvXPQ5bnktVwJnuA8blutfmFGgW3gB5f5eZTKzOOPajmXI2Jeol7Dmze%2F3vmjGm1xJe2HeSo8%2Fp48ojV57j6tS4ow1oixUOfPHIA%2Fm4KZsNJ331hcQMtl5EFldM%2BQT3Qj0BDCrmdw2dUI7KwM11RYKBV3mEKTbTcZFngU3V8peO165J0fTsVY7QHcYOC2aWrQBdFJPFSzPBDQRCno%2BuSN%2B7O9RVhgGcPg4xZLGdH1X6woFOWea5vbGt3qkn5rrEkn6riZuax0tfoahHdV3Wp43CYhW6nHDkYi0eTolHmSdzNY8EajAoD3ukgVGfmLP5U1%2Bojo6WD6joa9B%2BJL5%2Fm3zfPX3hP5ANtQBdxoX1BSk%2BWqw4ERk%2BXf0ARJazOYLpMke06eMCOi9BGhU9uTryjSqQr7IjNe6IK%2BvWYZIWA4igtsXqDFVkmIRfb3U9U2%2BBTubp14vk%2B5W48%2B74yKr86AFd%2B7WDF7hd30up7McS0syUWeYPPpQb2kre%2FwbIZ92Joa2ZgfYXEwGngXlqjx5SxcE00bZlJFLhVlsbhsPu%2BmF5%2B7d4KP27IIMASOzzQO479Mj7rNC9XWa0hNFrwh2U6ePVxJVml0VIqZI4WWW%2B2B40dX7lng7WGEKL0jVQi0QbnxxPKOuwKuz2p5NQQHTPVIroCYcutuxcd6Fpslp4yt4m9fKX%2BgboasqDkHBIXPjMvE89JtbzvNdzb7CrwtooRD1zZARbk2CytpIx3icDnqIgqHHt3zASgOKf2AzzyF3zze71GF%2BAMnBPjnkuKIAQ4xVWbe6mlu%2BM%2BFEEUddfzkpGQLOUKIztM61q15jF4no6jvtjVFWUEYlUirF9LkV0pS0zAiB6slkB80WLxVG3zT8ZphgxGVAE3HhuSk7vXoeerudaT1FkFXm1ZvUN6H2yt06Zj%2BHwaPvYqMknEhncaYIwXC2waaOe9oCWTEYqQaDB%2FP0g0%2BaDnDrYL7wyc%2F4aB9sTHwWRdodJXNyv%2BZvbbmGXWtdAf2jQn1lT5LMhmEGp6CK5o78ydYoTTV%2F3pBoiEtVqeTUbJx3a9YXR2XhvJoUqyAbPTRUwVoKCsBZbM6%2FbQJpqbzNNrLv1%2Bu397%2FDuFMKwHwC%2F1a1wLwgFNJHwxtT9Yqb%2FQU%2FMaXplP9KyNgfAUBFOUI0LQpuJjMEuwLVZPc5GvT%2F9XAWmanBezAukyg0L0ZiohGQiV6tOO1vYfI%2F5BCopEUCw22ValH4DjmXFTKCu148nmK31EN0IcuBMzc0j%2F07Ya32OvbQyafkRWeXqlkZHeSm0fkKpKrNCP0exPOxOXV3xL2PyoC42sEkyyv0hezS545HuVRm1CKOjGaFlxCVfvOo9Rd7CkjK4ED6qj2587cPynwb%2BSJFrQ%2BRZ9puzmaXksQPy%2BxlCXrNaLsOcX3E0W3gSDK727ZNtWitX%2BS3Hi88gad2GwlgsW6dqEgrghVJCaUq9UogSHfr8vX4VrcEOvTqm1hiW5X%2B9Zeb%2FGs6EJa8DObAWq0PUK78nAsG0eInfK%2F65XhgdR5DlgNZQjQ38PQRl04hT100%2FoEXlHOb7N08SAGmtBrpYF8MdhmJXo89nHgZq2hnUjBOPLCXfqnT5b7DIh5RkNwhf083ViXvtejab%2Fm7M675qh9YBq4%2Fo838qeJCJWUhpRgpOw4ZYM6%2BnfqBr2ozDYgS0npZddOXxSH8nLO23Oy3d70LzxnrfSUYVI7f0FMAgARjXI2RA2psPBwClun2Tlw%2FMfH5UlGrY2Tn9hl8KFFa51SRHZao7XhWH9OW4OweWNVP6%2FOOq2lS4u59B3IvwW40kd%2F4eB%2F5Jq9pDKYAUg%2BKnPnMh3z4Qg%2BSsgNeZFWSi4UvDA1TOvyAqH58A7bMxdnlJfc335Q9Mz7V7tM5MlJQYJ34ccUNR80Bf9y5nC1JlNORMe%2BxB2pdan7ZkeXVnV%2F%2BreGpBapO2QJ7JLgDXJPn5GW%2BTPvmHrFSwlZnJw%2FZrD715txkIAyutugo1h3FIHW3wG%2FRvkQSXprRP5j1ZZNCd%2FR4IGHubWCQpF2fNH4YrFk46aKcXpjbwGeoEvnc886CelD087OWfFxFeYYw09fLBBaJ9BEZrKBD0cyXIDZQyXVMEnmn92Hh%2FAW4pr1CZwoz7DyUWK5hMjyGiUXkodEoUFQWJLse6BcfNjPH4jR%2FYUJTyc5wKgfAXKHxNLWkgccymKSt7ou%2FOdZEkXuxaUZxZpQ%2BC87e%2FimEiYTFsrmZMfyYM33PqrNoeAtl%2Fu5rS3T%2F54398%2FuvP9ozzrAzAvT176oRMUMu2HP2udh30gfiwKxNgaJuvIc%2BVG8PTomnNtrIjOiO5MKXadhHTd67v1Jda%2BOsj9%2FtABEl%2FjCQvUcag%2Fa2fgGRi7QbteMRftUdxE5Hn2TavR2IXpuF5adNa5ec%2Bvff4MLqvVqOBAjZOn8U0p%2FX5rJ5wAX5L4ieuMxdStYyG6IKNJZ6Kj5PtfOdeysn6ihGxtUEkwv0QD4NdZS48VqRBX8ORry04Yk%2BPvscC9vLu2YkG2OIuuveGs%2Br%2BKArITQc4JTguQcrJmANj%2F0r6GtDgox0u4Qjpv3M5KMoG1XJW5et1O3tHJzQ%2Bj4SabMG5TAJySCLB1fWgy%2B5ye%2BKvFtMyUXX9JYR1eoOL9yK8g9V0cY1VozUJq7jttPviud6%2BDEuDIJh63E6s8N3WEEEwkbdow5uzGy4csLhtBWnm7sYx5tcMh%2B87hCI2WbTKoWpWhq5UUES5V4qTzMXI2cmVrbnPlrA%2FKPCEIpDbJDzKeRRES%2FyA7BurQRuYQfwGo%2FYZ5VTDfuJXag7f58lgx6I0SBsdcNKt5avnFHwm%2FOIDosledjFaIKAbD%2FA1SSjwOVIfg9z8EmvrdJhVLNWsxt0QpAAfJrwPbkNFl2Gkpv820D3gAMTfGzFBI%2FTB99lz8exrQ%2FgseRyj6gr0SRQg0DXy5pUgsadcTRR2V7%2FQGwT6uzS4Ppzqefk3AN1g9KTMypvRhKDAmU8n873D3Dg1VSEk7%2BSMja%2FVMsEUsDe816uNVjmCnmLU5BU18%2ByY%2F5Dk1Th9JMzj1%2Bj1DX1nQnYGMHlLxjr8V2cWeKI8C6Q09R49fBlev%2BGdeCJUpqTJ58c9ypLsOfsunfqoVBpN59BXwpT7nbAe4QDNCgZoktJgpmrFkCSxlDS9NtRIQMx2d1aeqbMBat50wxdZ2TSdDvIT9w377%2BoWwjHhcoCKymg83eN3wXXb%2B7lGUYEHqePLhXKmP6J0Z%2FCZ%2FfIqdGAWIKi5afW4AHVpGQBep4T32vgGgmRAVs%2FhiH%2BVFCtNlhhXDDTbyCjPiZiEyuPZhI%2FZdogrKjDY7HoAb7wwCQXd0LWbOcX8ntQuDxk3tv%2F51uHIRhpMSLqbyler972tH%2Bhv1nQi9tqiBRgwkaOeOtEWKHBWLfoyQbbP7GxCLyM2kubnyDQ0aUt%2FroQ3ly4bMLUK1gWTaJBiUECoHJiS711MdFt3KkifabYFwpLAnWNLj81MPGQka1q%2B0XrYHZl6lAh1uJQ8AT11RxBhuag9c%2FoVXll7bFiE0L%2F5CcAbTj0V9Uhy01LpzlrZdo5wb56N2AdwQsSyu%2FDnJvsiO%2BFt%2B%2BAPW8ntQ75hPNprAv9kZxcd1MUZrSl4cKHiYKsz3dskMh9EywOX1nqv3mhRiRw%2B%2BcuAzJlIo6GhxiapLO%2FAbxQv%2Bz0N1Qe7twew2sp%2BLOTSpVajphVwfZy9TGy0N5o9iVMrJ4SgyH9Y88%2BGNL%2FcNdtCSHb8SVEB%2BUDHcBsKviny1v0qML7YFLFmY5wuUDwrEnuuFFK%2Fj9iD2AKDVzY7xQdExpqKEvflfS%2Bel7nJ6c18oCX1S6GdROyYBreIXjvAbaOCn7kf6kxbloA1SFdZucWqnM4I5Z9rUUOcxBpdTWdEP7IWZSxTOAUPPZo%2FCyiFv%2F9BqoS3f7pnSqY9y71u%2BXDfhm%2FzctV2bO5UMwyP%2BRJ6H4oHP2BOHJLgdxDjo99BMVeODqq2DdBHmbacwrrt%2FBY6nQco3PMIHPdrwf3wJap8wT8%2FYLLPGcOBJfSiV%2BL2Ehn2yLa5IFXIpEWS6d%2FqOzzIVSAhFOrBUEOhnAvyEiSE4WKCJfuCPUcxa0KLIMC77uoO0t53z2eNcAQDLxoBHA2DFMb9CqpJRgSNu3m%2FGS9Kd6bk0DNqea9r8kYbJKb1curRNb04kzFnxBHKXQnQ1uqvriMcOA9lejitRpKI5K1WVFHeab1mqLR6yEP2m6PMzynHTA9Aa%2FXzT3JCDkk54PxB6gtlrPUrOpnQ9GqTQQGwCVAh4RX4%2Ffw2sA6W1Ixu7vyVwiYiupUsV%2B97VxU69D2mLd%2B21vi%2FILnjF3xH24GhQLosa%2BdML2h2Lm%2F8ToFCad5lbk0%2BNH72yuax1YiMLPVqm0JU5LzQIVf%2BLBQwS4FJkxQQdgwU20td27iXGmVgQL5qtAmrTPHvFbpbcfvkmpv68BMX4t9NVXdEkEsjuyTIkEEvINI6HER7dZp3PQv9wgJnfU4tzApCixlyEubxjhFqzR0LjaTkiRuL16QWeccSjgq9Q6WRGq1s3lOWi6A4hzHScKtwCMkZD8D6%2B1Fpt5fcSB%2FDo5RpPYzSk0HzFMgTWlBxr9jIg0lVR4Ez%2F3NkEvzxzEj%2FVPav9SVeqxuuQil1vtk9Ik%2FBqaKYhZT9IPk3hCh61%2B4O56VY05p9hn4cU4iud8eSY8ugJ5x61hOlZHtIPswQsE5%2Bjlu96Tl7DlAicsU5yvCQCpPLd0zXyGy%2FIBEwN%2BIpZUrGvIDt%2BQ39NJ8clp%2BsUf091na14tHnh3uIMs9xKNRS1vV4z5NqcPIUDG8304fVjaUkHyfrfZN9s%2BE0Abn%2FAydVpy9jhuwpBOY7plqoDikPfjX9l%2BCXfdq7zsrIoBCjGGjp2jzgF57IwoeVVtQg1f1CkHUPnxFx%2BPPJ3xa%2FMhdwJ7r%2Ba%2Bug7wM9zy2edPGZaR52tsDBwX2QS5pShrto0cMg7HlOwvsrStpN24WPSnv%2FJ%2FIM7gxx%2Fi9RQOzRs%2BsuD%2BBZpe8bpcu1AHIHS79NhcFzqEsN1UVAg2tSRnpj6QIfuWc5m%2F5ox0CRLh74wO8Bbcmvb0eB6ilUZvhF3dgMuwA7Iz60nJFbpsfLtAQqbv2jJg8s70594yVowTGIAarRyihxNRp7zmHu1dP5XevTUsXcns9Brq5Lf2tzx65NodVF3E%2BWSRZQgrkMqfW8lLwUJAWRUPlawAjLjcEPjkR4fUNMQknw2xSece21daP%2FWLf3sZ8AZSq4OmLwAF5VQG36qUkM%2BZwaFiMLfaGqX7Tszl3iT5BDwEsxvCyXNxYqFDV3tYWWuW%2FzA%2F9L%2BcHkymJgQ9l62oGfJxxS6Lqw1yL7hx8zGSVgv%2BMMfxoeTE9okBZbD9ujwrU2165GJrj3z4B9%2BaZLXPT2u4N3IYd1naT%2F%2FlMArujQFwfErsdSXZttw9Cx06Rjqr8E%2BLcg9dzFXP%2FLBICNlvlIQ3%2FF1cx30DnhneiLStKQZPEggroJUHFVi%2FemDN2iepIjzCQIsHhRlqTSG2sxlvuY%2B8MWt75v0YHT1x7yIRVgPIQfoeP8pmV7cdN4ZdoZQkSukSbfGRbH0GT5%2BElfy8hHMeKtytHQrcEtpGJ8H%2FvvOxWfA4AY2b%2BRR1A9fp3pkzvTLLdrdpB8kd4UzZ%2F7UphldcrN4PT%2FjBpgDnbdbCi3Tc7ilaXkIVz7jfkudRLgJeEQkOKBYT05pifH%2FR708xuvJ4uoiaf0GLkVYoNH2of02ldn%2Bgqq6eBldY%2B1VQbj49KPorhIE7rXLdlMYxWG2lZVWRJ%2BjNrNmy3MyMCr%2F5XRNpRcgR470XGs5kmeHhOMAua90uP0GciMwWK0zLym5a27Fcry323%2Fk03ZXF%2BH0xjC6U65eclPjpi6HCl3QenBkOuAUi3b7IMAq7N3C3HLrU9%2BlKpu3cVPL8Bd1lbEr5T2HVOSho%2FD6Y2DxzpdnCNYNb9RgfMkLt%2F0bjrwCS0LPjG%2Fgm3kuK%2BQwkZYZgTU%2BHUiW3tsTUMXmw4a7%2BInI05AIX7XoV38q%2B3x%2FNxE6EuEMB1cQDTSAFB5jwYQYpxg6E%2BztgLCEOUdUK9X4q2jFnYmYVHsl9EUOd1NiUr7a6M6QL0l6SPgMpHHpke%2BtdVRNE4Z%2F2mUcz9YLzUtdCgUgIUo7ZDnSehaQ%2FCC%2Fw6G0ktFPgw5bSe0%2F%2BTy0%2FoyvdAvm4i7BmheGOvU0V%2BHRYx4NiNKJ4Wrzgav6OUb5qwIgQ6wksF4zPdKQtKFiMXvODyn5ATZUATgyUQf4Ejrp%2BfIQkWY%2FKt%2FX2wMdDPmT0TPhSMKW3OIrgY9%2BMSRxlbg%2BxQF7qLP1wB9Vi2e4M04hXlHWwq0N91xICa%2BzM1bxxx7O7h2l9fQZMeCZkmGGCGDcG7WXmYyf%2Blue%2BMTuXestGJAxBzeigph44N1BtEIdpCNJE5wVpiiNMhSn%2FtEGfkBSUJx%2BrrAeBKKdWo6%2B8mxo1VCIvWsXTpTNkvxh7mzGhvgEY%2Bc6S73UkraLgLnL0BjZVuEBEYLhz0wLkw7WJT6c9Tf4gz6gUr0OcyrK1VnWZba2HgGSIzFZLgjOoVwbmNl8VA4fziStCuOzxUXq34RReOUn9dR8jMWFAjyzxud0vAUgEpC32qjI3EuWQ%2B6xb8k6P5N3K04TW3XbTqE%2FxNuSCwDS9vJBehNH8BvmmusT1Pnx%2B2kRCi%2BKee86LsBD5daJ9G1iXBx2As1kOaiFH%2FvL2zyrxOlxmW%2FdsKs%2BJlNZ%2BD05MzkQwPd1iQdRbbRiyCZVxsUuE9wSDWkmlaM5dwnDtTpint%2BA8lZA0Ecvv3nDOWt61XONvWlhMMVzxodeQ0kxN6tGIQVH2ZNVJmK3jCNl1pxsRe2EoVP13wCpYtJSp8W4UXjROfIrHimQg7U9f%2B%2FfiFrFUrVKIUPrAL7ouedc9SDkwWspmx1VT53pD91nzhDxrG2EJ97YJWQW07HIrZG1pu3g%2Ba7LaQUuE7eDjJt5UxAqdgUtazty0FkyeOAXLgqhGqzQkRwvwCUwubKQg16IwkHmfgn5ah7fLlhJ6DZMDASE7vA5z1CEGsAJ9vyDQWNghlfXRMrkdiYYS9s34leVA7LN8d3HQk0bjOtzOc1wpQEdONPgiAYnoKTCVh3GWJGMwyVc6VxvdaOxEtAZotxBUR7LdKVcWbqv0DquwaeNFJoXYUXOZFmqeDqwDjK0AAPo4WbsA7G6t4T4KNP%2Bsn7FF%2BY5r5%2FsBrmbEfdUHPrkRJPaqlH0Pycm6Mdw9i%2FT719M78y7ztIfOcphNyT555H4Q3n4g8aLojYIuLGFJUwx9%2BnVsgPD6HrNsTaajdIPoFKM4XkP0rFIGRW%2B%2FlNIpsQYLvEusQo1xVela0H7%2Bh6aSWfa%2F8deHmqBReC5POuRu0X%2BZeeZIbLY%2BSOojbS%2FNpBtc80Z3LpPqGhnKgWxC8yW%2BIetkuqfa3GAfQhaomrq2AjRdqWweDNPPenTUFJ%2BxYmQVstZTkMRJf3%2BlN5GyGhmp2NJB1Uo5%2BGOiVAUDDYdgbRHZZXz9%2BfazerxE3nfwQM5p2%2F9UpWJ02WCVk45kupbS3jVbZAJMcYVAru7snNqM5c9F1RWEu5Fd21K%2Fc7W%2FiVg7Zr0Fwic%2BxTa95P5JcnzyAzXJ7%2FEj2YnF9AaMufg5aTmMiOAz4CMothS%2FFzKnqapViBIrib%2FhMJiM6%2Bg0qXkhf%2Fnftx4kQNxM353UjMzhVf7kiFkPV59CwVRXfYucfgF8ZDwbWqXa29wN8YbbF9c7Wx%2Fr%2BJOhwUq7m550KccK6Gg36z50vE3m7AK2YzFJt%2B4gLjh4U9PpK4iXB88IIJbYa0FczJbLOeeCMhbCae32juck6DdUAPx68ltHoAXYkQ0hs5WhpGklKJQXVCY7ufIiW2GayQT3nY0ZELpTibDXPTE%2Fy57mIFoju%2FBI4doISeuI0hLkAi9K80ctr3tf0W%2FsxvebrqwH3uBQ21Ini1YJpunlfwPp1hbCjmpZ3Gxt%2FZ05jo19tt6CGkK3C%2FcXSSHVFgbroDjJerJ7LD1VVyzuFXtg9o0WbQjzT%2F893cR8YnsmRsg%2F7h%2BrVVOlC5VuE5LcWoj6iGC5yCDRtoLgreNAQIBb03zgEEnGOKY4A4Dp7cXj6lZn%2BnhqtCxmTeoOpqrwN1GnuDnyvLcaQzAAUHubqCuQI0kctoDFDnrh4yRYSa7lzqtSJyTwih%2F2Vqw2kR1BaKohaQaj8V35DKGsnj2FxaGO5LVW6GUOO%2FuMUK4EX%2FAj%2BbBO%2FN2RyO%2B9QlftyNdsviuykxMdP4W8gQOMvjSMJnTHIk0JIV6Wh8BRQbywgO3XV0N8Hk%2B8zKr407U2LYzmPs%2BHcv7IOXy2wu0xrEGauB3ecZEtc8TZY96D5Il%2F2%2BmLxPMc3K0j2sMZDOXhByPWNNGnVVVfijVbQd526hPcky0YnL0%2FSlEjcUHPcQm%2Fm6tlpAmCkgB6kawqfYDoCkIFEMPu6H05igQN1afJTpdGR6F5KYQvSsxCbxia1cvd5QNR2KZtdeKUtUw2DmVstMBv3DTYup5s36mi8FTx1u85iLmJ%2BsnUcOrAAgiY%2FuBxfKduesJTGctxEU1Cl%2FdGXGKdj0IePH8KsQTpiq4cpIvcaiKAfZxs2nwupp6NvdC6F5aZOqdCKisLsi%2Fl62%2BDOtwizaiyIxqYJgaeJMvclekYdoh1wuh%2FTMyrZO%2ByeWYfPLRtI10tDMgxlivJtgkIttWETQLWnXUv%2BVcXQse97%2BhN4g%2BQp%2BtxKxJP6hpVVjzwXJWFF1iP6WTFKX7npKKkUYk7OiXpXt71xLqvOoOP1SgSxYtKShwBV7WgA1NrRwka48gpqGVztOYT3ESuCL%2FKxkO4tu25iJhqR2dQcvenWmZ%2FxdfiNqOpZHNv2lw%2FYw063hjanLXMfHIBPjXsvB5N3oxDeptMGDH1%2FZXhWslMJ4Lop84Kl1wIcdoFJ9GxOjLt2EPO470M9IlJayk9wRI566EluUq0VY9Md42FmDVsXfU9de5DKsAj0MkEjvG3EoAWBqj4C97IWf09U7rW1ZTGAq71XDnUTR%2Bd1aWb9D6hFeFqCq9gpyd3ojiRkH4GetpqFftPol0LRu6BgSD0UiPbiLpzTMSVmcbqniBFDE4UNsDrgTqeDSll7dNmYaIljHOxrE%2FU%2BVRfT4P3UeuLq5fRfiZYI7bYxMSsmZ50H6z1n%2BqYbzPNoelM3NN0nvbEv2Tkw24ZXmREKQkDkwiol%2FcOIdZgk6J5ht5EtWHFCRHAko%2FRYWpK3KP0PHnV3EfuftQg%2F4VRI%2FAgDjGSO%2FDkE9pMMczOMSbJnhebc37Spj%2FBhhoDd1Ox8psD6T%2F%2BRfR2AX5aAdCyLuJLvSVh1HOU4S9M%2BIPHoh9wXx6ueewDjSgl5jYIMvHmjCU4lUtAEuxT7gnw%2Bb2CfTEDBkRmBseNBztIPlZelp1Nr%2BrEYKj%2B3BVypdDuJjHs5KbE6IfgDpUmjPkVBJT3NIKSUPAT5HIgxrtZMcB4Rdo0D08mdphQueLnMrlxMz%2FSPFwponrYdbOxmfFJXkR7KX99Ed4qeXuGdwEkCZ%2F75EnSF6ngNqDqW8vEWXj%2FLWS3uKPcRXxiqgQtSMzzZ2gEJD1rHhzKOrxo65fggrFYreLyJVWaOQQLLW88zbkgBMp%2BTjZKfXI%2F7lvBGB4XuZ2x7PCAfevncRnqUC7M6I%2BCsZ3j1gR4EwozRc0Y4UF0GYS2jo%2Bu7aJjXoJsaFkTDRRIexxIg6rO3wLRf0mnOCdcZCzQiQ%2F0BV5o3nkHafInuCkGV79qRLf1%2Bo%2B%2BHpB%2BI4bnWD%2FgtCqT8R74ZSc2vCx%2FSThEBmA%2BljHvgULQQwwSN6OcdGDfvvPpoYSkSwCUHlcWJoS8ljh9h7gnzWJECWTvKku%2FMrwk1oXPj39BDHNi5hlnAzBJV0yAT%2FOkdBxz3yj9sbwig9VAg%2FrGp31DpHlhCyVxo2mB4wW%2FKKChyzlZsAqchZJ%2BOs7NFSUoVbvmkxUYO%2Byk%2F%2FI7lPeCpPlOWDjXhnRx4LyS03SFqVAvE9sUI%2FA%2FGfDOlllX1BC6jJGN3nYh57Vp2gkca55uNfAQTJ5bd1%2FWu52crlAU1x7NHX5Y4OB3Z0l4sR2Nisfgc12RvbZTPPiQTwGrnWhWz7tP7xmaDNnj3IZMLgWDNXMZoTIcX0crsNoQgDYdgeUm8hWsNZ1n4fVMzF7JeIy5DyMplZvh9aRNvgjeMCwwdkgNDs1gt39XtWsFIylHZbDDTB70ysBM8aiXZ2bnXtuanz9i7uUxikhxEjrX9a%2FVnF5IcrDe9m5ksIABFR6uAiTQAWDWZ8z2%2Fw87%2F3bX76%2BKBGdKXcdBk8SlB9FSWYC1n7eqnXJIZtaRn9HYXWQMT4D%2Flf6NGoW09d7E7zpNBKHQT3hEY4VYUhqM3V%2BeBEWNUCWlfs5uJKpQl%2FyMfePrEyDs0%2Be7TC4aD13aTM9MqGsUnu4%2FmAG1xbHzwaASc3CRgkJOh7vNBZSXnCqtD0OUBNhTAj1NXubkbZ0V4CmSoTXwDEwRgcxM%2B1S4kH9KlvqSGkBNn%2BmYmwyde0Xfic%2BRX%2FVmdiKU80gIzteg7Bn4LdG2K1uFcupN1V7Taw2bdM1n%2FGVHKRkPkG06FSlAuyW5kxx%2FI70MHTcYqW5l3BdQKX5gY%2F6CooP8t%2F%2FkpnNF27d1QsuxK0olDlVFe8r9leSGsqq1uzml1WoUE2ozt1gq6RKv3%2F3FbXVMuWksP8btITfQOgdGVch0n%2BGs5sPRPa366ay29%2BuGxazwUiUjUngJ52oKEIYf73ue4RzaGJX3k1s4f4GZpe7EeuPjlCJjvfCCTN9AfJcpy7yuFNBhWaPvfhS6GHwTWzdhiu5btW3RBuUsP0rwaEQ4%2BC%2FRl%2BzQAJDC93O0S8dnYrgXhqD9Ieh%2B2bNZJG0kpzfHhgoGiJtkZDGfyUOfUXVRgOW3rf5BBGlnRM%2FA0CcY1TD5IZxFB3CbiAj9v1R0rE2k0e5yR3W2of3S0668FfIzofkbHpLFe1Adf%2B3V3Qr2AmRq9AZUS7lGae3mKSzOK8eRzqbabFmQvswuRJDg%2FpHGLO92jptJaj3HDYvnjPhPL%2BCMV6cCrE9ezXGu0vkwi74R9E2xkg5AZF7yP8cx1ssy1H7MPJ8AncXBlWR8Rv%2B%2BvEan3hIpFTSkZ9KsDIuA8CqLa8%2ByVDV56SaPakpFFs%2FrcsFi8E1fcXYti7qKKw3CKoubBvohdJPLq2BPILRlg5WchR2LVGjrEHpWHfN5NfZFrf9OKwnUV9b6KHoOmqubMXPP5OnAldXhle5pljsCUZlLgr9PIMqq3U6fBAl1AhUNOXGub5covJUiMlymcMbtA84%2BTHjZrePH32bK4YB66BpM7KKX7S0lPul8P2IeX2zkAcnnEINKu0eaeVhqouPL35pkSiLIzXjS1TKk%2Bu6VibKmt41vD2DMVR4OYCciwNAWd1Uckgk4Yvn69hAWAPtw25qS6%2FFUDwE10JNpaZ%2FyTdJASKhHEs3W%2BRIvHoyBUu3g8KgU8aLF%2BS69rX6eP2W0AKOgpEdEkvIHgH9u9H5F3NWZqP01ET8Zdd6Zmlq4j1saf4nddBEKFHBgwbK6OnyF4SnpuMjEmecBEoxSAvUCZk%2BnQSNIQcNg60AUGo0JDFd293FUlZdZXX%2FclCS7aN6rq1%2FQSMROFNl%2BHl5j0uiTe6CgyGB2Kj6Vlzfmz4bUB7Ur2OXOOVW1t3jlVwXL3nVJ0e8Cs%2Fj7RwwnntzadlhQlAdsdhLpI%2FgU%2BzYJatMDrvgFJeN7uOM6e55bpeBN2Js2AhysULglL55B18cbKJQdOx6ijriMpZa8%2BnpqeUZgRl%2F%2BegVryGvnpwLIyBbx11%2BCovCTLz0SK%2B21hnrv0O%2FFzO5fYtTsYc%2BL1Kd9t2WhnFU4V4B930uuwYfYHiF85wnnP1X61PznLQso5wCtdz5rHpXkiNg6zYEKlB%2FZEk3%2F0%2F9%2BqBQld%2Fyp1xKd9c%2B6JvZZ%2Bh8KtkuaJawA%2BGKwsMtpHyLJwg%2BtPYtkMsPinCxzUduYSptwn3ciqNs%2BHTVtgSQnJN0gjAz9Mh%2Ba9cY%2F56YC1TqTfueSObU06OdQQfZCNWzvbWNOTTdldvQSufohxK8TaETyciKaeuueKZxaYUJnotWXWyB4mt9z%2BonqaysKgRVcTOGEFJf50Wv0sEM94KC23SM%2FvedLCXI0JeY%2BsJqVNX%2BDkb9awXGdC11a6bLkuHt5V2HUY597fvS6JzyX4GXdSgEK2WAtA420U3NjP5YtTOIQ6wneYiUHRhYaTXece0h4eGjwHoqAoSK0UF25x96P4QT%2FwAaq1%2ByBRKJJu3kQrwW%2FPcHeP%2FU3e64zeJhpEYNVLJc4ujQBA20k%2BjDlICU8EKzd44a1msb3q7Zp8pE0TOFzQ9nLUqIMp0mG6H8wo%2FWiRogrbcOMJQnqyzI8ry9aHgg5g%2BlPDb4JWY8Fh8nNAvsLAE7WQ6KG7dgCkjbQh6s9K2pfsvMx8FyN0Y48j7K8lCxwN1peB5y63LymNmkUDrqPrw%2FeTlw8FNmTACsaTpve3Edch%2FMoZMiDF4ZpSZIx3uv%2BlAdMWYeh0c%2FnUHxWurt8%2BgDPqtCySJ%2F5zlewisPpWb8WQXGzfoneIIWuf3opMYhi8H%2FQH3nP5%2BuYMaOL2oGbDC6NyMvF6QkRGjAeGKM1F91wCCTbnrN9LJmTeI7fUOkHMwcVw2gfYc%2Ft2aQx4Mr%2F6U5pJgyl5j1OUvHCOdDKmsVTXHvI9VTulNX1EEPg8AmNcGfegRwQRHRBTYp0B1bnbq7Kgr9Temn%2FbELdbX1TqUxyJA88SBMCH9b1pYClRvxp%2BvT7XvvKi%2B7suZabFvT70I8Jhu%2Ba0Jz0EwlzrV1tWC0nkVjRRNG4teTuo%2BUAYHYl6DXz%2B%2BPwFCuNntlW3cxuNCKDbnekb2CLzcUCvGoVpsQEuprA6P50F50rHNPkPpqZe%2BII4%2FQyTWQXRbDPHC1eJkAmccKTSLWJUQAl0SR3uq7dDby8PLxFuVfGi%2FyCcfVnJF4MoLHESIepRT6aB38q%2FjZ5pMMQW%2BzCe7N4Tp3HE1TFhXTExjKpTtgf3cbMjP8SrEHnHd8i83YcbJ5mqaVWUwlu%2FFTPgmS82cyTcJtJhAnkLX0UBmylNuoAEkLvW6noB72wTHjztFo0w7fINKGDmjLmIe%2BshZlwhTu7jqRAOK2X8BdyiSyoLoh0uuXLiamlG5fOFnejr96rEb4c39ZHeJslYyuXlMD8kllMihXt2ZGDSPrnoy6Vjs3v69o61gf4xl6zTfv4OwDdfcY1E%2FxcVMfaHHtY9b2dDXitGaJbScg1cChaMKkE7t653%2BnVpr5uzQgvnqAKSxwS1SLoNeHzwZI2%2FPagnpARBPkfwhr%2FoADhn4gdxJff1w3fAhJTgDQWXCKksBFi9P6JGMsnI%2FiFjxZz34qNqnMYofOqj9DnOjsd37iYH%2FWb86ttvnEbjDfy3p%2BHm%2FKRwc66U2Gldw8%2B62sM68Ik70sONJsOxiystFkWtuqerQW06CSUJPFxIkKohCBibxHW7hsoQoE5WrZVT2VqWsxZPGNFtiBaPu6q3xA3ZVnnd4YLTt8ghs1zworGCAOeeQNeQW6Hq02iyvl2XU%2BFFM9fvvtWI826y7aeJTDpLXKUXgfU%2FVOzB3ZalY6KJBxb9vIH3eL4ngOwDGFT4Dtw9eOW0AT1nqUEFAHw23ABI8BHsWRDtno81bhOrSEUqPwmsbSThPt%2BST6QT7OL4XL1mJ6vdpI0PPm5BzAiUBMEggaxZ3ALM3d3ld9jEEYpTaadiRPK8SMvaXAAVqrOpmowpTeo0qBTOvWAaRAFad69uEK%2B34ZK9449JwsULGQpvzt9dLzNRvGmyGIgw88RJixdQSv%2B6XlJPJ7ueqORoM53Q6uIFQwmVjXeJArE%2BA1bF7S%2Byj16sRW3ovw84j9s3Rlx%2F2Mq6nW0EaPIkYV3BfI59wQ5Ac2YVC%2FRx10f5HDMbU%2BJ%2ByKY8xfFUZt%2BV1Jps44DUTWDl%2BA0WrsZI9Tc%2BboRruR6QH59yNTSMOIi72YzaKwpNY0em62t%2Fx67R26dVyU1mvFlGTrvElqZ7HTthR1LZohLRTP%2FteBi96QPhsgqli3dsj1rizq931f4LOYJIkirY41eT6oH6oQ4Dk6zq57U70m%2FSWjjDanhEsufm1dI50mwUTnW9w5tLB4fgu7vuhxySwiOSlo7NRCUSlxwZshKeRu8K%2BgwmZdDZticyDtsg1v00JBCLplJSyQTIa7Wjm%2Fj9VT3Am6DrXqQz4ftWYb6Y5d4A2fctxViaeisMiI1ScZtetuDblan788LLXaXne6fb4zssvrumbo8r27A0sWv3PQ%2FgN5xlLAgF5iPzJW6K2mUsrfTPp0lqIJ9wVgLt%2B4eh4jAy9Lr6oKQrs5iiPrjPb%2BZX57ZcCvLC47S1w%2F%2F3NAdvEkHgCUwFwGR6hR0VVh3N8udGLQ0zbrohyglhYhUa20bxrz0eY9NoGy%2F2k%2B06FDjhPwWm1bErJCVumc85Wy6qzRJvT%2BMg%2F6jCimXWRap6tyqr0ggEsXGVzvedOY6rzt0LXNv9S7jihmSy3fUScd%2FYs7m%2F0POPzQziR%2FD7L2tnnhuIB0YpXaMr4IDo%2FUM05%2BxT75bSKap4%2FV5mn9RoqzzpxMJZ%2BXb3Fye%2BBwiR4D0XLmkgyAf2VlHloP4Xm%2FyveBbhFjH7kk3OfgvmGfMCAUiGIjYICmX5vyWhZ665vwHuXqz%2BRij4V56%2BgdKEL5u5z3US2d3qBG1BVGELfaXWoEF79obhTUvHTPg0I5cGcmmsMHAQgWleC0hMRsLFonryRayD2ahfMXTDN%2FPm6a4yDQSsBYR8C4BvKu%2FBKcw7iz5LUfA9yhKH0ZwQFeI9oHcKUSbzAkLRe6Oc8WBmkZsgOCdzaRt4DonVtqgccquKQ%2FdBzr6jTTjKjbkRy9KA8LdzcwHlcyHxDCTrLBRXeGjiL0aEwdGvQ98zDeXGNmkI80WsaWAdblHvyQzTxyHyBZrp6l6KSfrpTaqp3qsn%2Bwda%2BhTdFksPnLQ3dmCNVnRd6BBy3eTuWRXDMUSzOFyn%2Ba5pc0m%2BzYzBESwFNr6ZW14k%2FWveP99unOm102Fr8qjYMEA5cOU5J%2BbkNHerPuMnlXY50%2FSGwLPUQO7JfFLjRav3ZJDODq1Ei09LP5JMese3hHjtOQqobD6Khj0piDfM9vOEZV0S7HqdTRZYMyTOu9Bwv7LjNkrYMa0kGO9Gbb8KbSH%2F4RtlJeQxQlxi6WLa7Z7xdaBDgZhIVV89TEjWVzE%2BQrKehb2RdXbIOg6ArJDe1SZF4E4lruqXV6L6zL2tAKq%2FVImyHT5RonUZIh6mG6SZkncvu8xSQsRvJQ%2FeWotVoB86oeyFU5oMCjNIjocj7QybcEd6n3q2vWPV1bpjbjLR53l2EU3Oq8WwTZNQRzEuij825SCnqvFqNJqttdE5A%2B%2FyVl048XQtyjSN4LbSp3EkllfedStSO5a7aqRMbJCdc%2BSVfKtDWgE8na5nJHraOgcLbGL%2B2vh%2Fx65CbfoVU6Bwb3%2FN2RnyxNGa7uzJ1JA3OL1377WwcXf1KixnNRDAFMWpVAOQGhzevdwVSMs0BVhoteP5pPl%2B3s12wn71kY%2F6kqrBGEPYkxVIlotvrON3SRKTHCa1nIxoQGAo%2FOegBSzkEKrY92H%2BligzUtcKIiX5%2BBeZolsTd3GP7ruya3tqz98TxWZ2RfQ3Xil906i19d%2BqISxKydUQIVbpBmlXFnhuxQATLfbC3q3KM1k3gsV2EaO6v2VbzzNwtfV5C05QCK%2BxgKEpC%2F9eL58ZsTr4vWCwliMD0fxtRvNaCq%2BbZaj7DqOtwTysAiQCCR86DQIX5hdpQMpvNKBZPHfglN0HXLNRxdNEcHacskRzb0%2F%2FsAel9Mx5yHAa4LWwhzfUo3uz4xhkRidZhfn3fFiyrld6BkjzK6O9ZbBG%2FoOn%2BKFDbU%2BIBAIzXE67Txb5FTv5UlETbFxBvYswoq8kF0yat4kSSYJpuQTT%2FgOw%2FWsSD4KA7kQrJKuxGS4XHWDJPNvulUHDbzi90tACtyljM%2FG7n0NGWHZ7KAnogYYt8WgSMVtVR8C5dH%2FDQfqx%2BwtGmQ%2FujVQIBXYUsyv%2BNJftWWwkn0uEPVteiGycOrV2CXR2ZNSC%2BBXbaulIp8tyyb%2FzZ3YvG%2FJ0XHh1JwtKiVUy1OvGu7D%2BZe8ItqIydg4pksAflnIo1iOs0C0H8RaWzqpIyoVAhImjcfnhNGJf%2BKAbSFB1IcR4ROKv0ncs01nYOXkhGWn2Xdy9unsMZgGQvUZx7IS2WqUIIVDnz4CwMikwv7QBsN0LeXdakUn5wX%2B7q%2FMsZIWkrhsuXctFRoRnjP36wXAvaQE38O4S71yYpKuuTPtPYVT9%2BZaOaQ%2F%2BvjDDaZmjT3lpwj4S1bV3uWQQM5dZ2XytNisbzHE56n1B2YmDq8fNOUp2Bq6ce0y5vONnnFlKGxaEV6qRhV2QpTXh6KZ3K81rylZ5A%2Fyvz7lho%2FvIVSfBATr9flgq593RRceJZ8RKwWLFCYEQL6Fz8FWxrfKsj4fNaTdMq0F5%2FeOM101zEsipxxmQeGUfpd9YXyR1xJ9HLese43kbfezIlXnNlUuCIslKbhNBlBNKFZ51Y%2F5GY9TmO6Kuk0IWap3eaBD2ddIL8EslkbDfxuO3hh4kgR1gv%2BfldLNOuXEHQ%2FvOsDmihLlKe0GEBneaNAAzpyn5jTrE7SCJmxn1TelqV60lshDRLLbO9DD6alsvNMtJ1%2FFnZbwcsj%2B%2BjCSq5cd3Q4f%2BakBh163psM0Jr5XoLSXZptlmWSEzXzoksAQZtIP12W5NPCxmkPOvjevozpvvaooQK3IEqGSAD%2FyWIO9ms5vK9q1Tgo14t4WxnJ9pTV3rXdcrzDXJmNCfWpTehacIh3w7n9bmzRjI5dSmgFnajkeChBqZji%2FVtg8NyU2lpZA41FkM5vx7Eg5x0ErA3HHw7%2BQYWDfQhrESVBBt%2BeMHewW37lbHAXwrwy2naaZ1PJlGhuzRxLmS9aSbpE%2FdxPoc1rTcAEx6RgPP3XpGAKqsJE6OUNJJI2ganDPK5064edm2aFyo9WVeCyCD6tX8OH8oWFrhFaOSzMEChdJgZzZbzHVIoIQ9tN16%2B%2F0OeA9rl3Yfit%2Btco5rzRdSPdEER85%2FJB0ExgOdvTmgMkZgrxyCRxbK8lm2XKjPD6eeRxAhGxBufcEa0EwBUKb1tSXc4mRdZMxIBXIQRZAneZkUh8%2Bj4AuBthjlyJb2BuvLw81rSL0qbtrBjZC0Vk8l3xmULjaphJHj03Eb76HO1MVw1jAYeZM%2BGu8Uiq5ugaLV2fHLH7IVJKwbqpnmIsLa6OeTMbDcdjMYDSRl10Xt2RyWWr9DQ04arTtaVtX%2B4BHQAeKMSn936XMFfXq0d%2FHCTU%2FITFtWVm0ohDGfZOr9DHQXO9Y4%2Fw%2B4tMkAeWuYE0m9SvzXWawr1Sw9IkPpIR6hqUFZjeMQJ4VmK%2FkU6SPHmNY7xmkxOjxqGVYZZa0lEZUzUCD3pEjwEY8qI6t%2BfFymHEHrX%2BeE%2FRWFSRJjo5QbMmza0pBN75yViNKwnVvvoL7mv0BP7V2vinJV1oCG7epiisVYKZDfXXtwGtVi0UDn1Kv46zBBhov9DfjaLrplcxmJ%2Bn0nOnXJZgiNXP2uhfMpWoT75pQfM8N4jZQEEE6WV3jWycT9Im48AXoRx%2FGrqe5ZrKI%2FnOZ3po8zqNfZFYPI53m89QlT7Zd57p6PgmYJK4VNvTHZFBbWrXHrALzaLq8x5ZJ2WMcd6gWEPze2dwuuE%2BIRHi6HARBmfkJKZAVb83y%2FnmZsac3Pi7UBliMlURkafc%2BcqDAuGl77QIa9pK6UnWNhUbuY4PgrtQnpNyFx6jRRD0hiYG6gNkGO%2F606paemqx%2B9g0m22Vjg7zp5oOM8tCv9hC34lsnIAaZ7wUY5g3vjKWIL4azxqfhToaX%2F7FNC%2FFFhICdpZ50jmIhs6uvv%2BPhq68xO%2BWSu17k%2BX%2BHySUUzxRxcca%2F1vFhVWCOWHTJJWJUhTkpZExAWn3AibBYRfDw7vXkIt3uYqhakfXYdbQpWSrUBVwvNXvhmOw5pGHFmkF7V7WDMpY%2Fhz2EwxNv8VSBUWawj%2F%2FUdCijBCE9Zx4FJ44S97JDTs5Jw6oQz7ugsJcHuOEl8gjaN6qT5qUu%2FfzEw%2FEdyAxAcSDxR7D5gBSBq7c3Aq5nyAvIHHALak0A38zrEbJdnoPmMOCD%2FELm12uxItF1KyyC679gnMPRvSy0L5pD6OMIHJ%2BBYi%2BQdKZGKrYSmMOdZdob0x6l35HlXbIFxlcCDmUQq4zq82cn4UA0EmF1m0NtErzoc5fNeXiOzvpdW848ejL0BHI9KpS4GvuTyUNAqbCbkTfmvETZpkPSnHtqAFmz9myJji3YgIVkBwtWovtmFl5NJFHcCd%2BHjm3A1qZbMf%2FHCBPvIoie3b0FmIQovz%2Bzn08z2piAHaBo4jSymxHVNlXanzpf2zMH0%2BtVQoomRkjzImRGHzqZt0lyHrPiyYYoE6nJj%2F%2FVD5NiKTmljLP67i4bsbbKzQcMydJQrutybyLGYbF30yRUPG%2FbTUgdxfXwdZSfmgcUWrlCtfachgOHTzwVwlnIV7JqNv540ch1YVqMtABdaVr70u8ELjhqGNLt5yL6d5ek5jMioziSBElZ99pNcCeSo5ACc4yTCUHcyTZz%2F%2F4u4FPehvRPK9zZpxIZjLEeFKmydinhbu9PJ2ZMISfooXEVaStv3HbaYol7iYVN6rZzcpt0Vhaho%2BMP%2BaKUXeBlPq%2BYIm5PmoNbRKe0PW6HWR%2FHA8pbieNdrUtVm2a6wcZ%2BCiRZzzdaH%2B7EqqDvspfmKFraa3Ysi8XT0xICk55eLwnsjEtIW%2FOSt5S5HMryvphn7UotYVQMsEFEq48UKnq3mDbJkZ3kPr%2BCbgWjEa%2F3lfoZcHxjhNNQO02sddtWR987%2BJiMlr9pNe3CjQ165g7YvFKuqva4%2B%2Bn2Er7AosTC%2BQQhQezpZWI35yKwxWPnMFK1LYoM5Jda1qAgd%2FNr0HIFzPFNf7MJS6O0bSanxs8JAuglQQuPN1IHVD3dUm6yi%2BA1QK3V1ou3L9OKG8Sfz7LhRiNW%2Bn7DZMkqO28%2FgickG74Q%2F8Yea8th9QEelzyysEzBnvVHN2pswyNVNT%2FK1Mdn2bzfCUBpBJSVrVaSPxx8%2FdvypuHzpuIVt6kSYoPEWoYxjrNrPohpF%2B3mN8P7JvqFJJFWVykbM2AHl1FF%2BmK2DSHPLD9S0YfhjWMBvjTxoblMpIYwguvxHTWOqdwWTx0tEBMEbWYK%2FNoPkAIHNhqhDm1ZGuLOITVSYItVU1rmaLHi7ziPCTmTe2w7nQfUxsoiU0MXWP4V1hi32sTlAmOjOmQLfQ52sFFz22wyDEJn9uHcZCVWOkIIno9Je6ukrW8KMN6mr7R9dE%2B37UTeJPBt0C3cPL%2BIIf8in%2B%2F9bdiv1yV4CoITuc1kjl8iEzpWzN1j9Gx9kYzSGAmPMYKSAnFg3go6%2BLLgKvzDeiiShMIR0sqF6xyvx9y0ML4x6u1PDkiMA5qC0l1umYMhAkvU9XbdqWC%2Blz9mmgI9tt9cfLaYvDVhpmZBtJqqBhgPU8W3tcJ7w3JZYj%2FGQ8oGLWKe%2B7wx%2BXkKQkFIMOAk8fusFOUx87AeIsMLqZISniyzQWPNlOQgtr9scGKcBNU5snwDd01trtQYwEiPtrvRsQ%2Fjrlkft1RclRFrveLb%2F4CbqByI3NPaSFvYV6N%2FzPcTqEIgIVBjPl7nx0QZ%2FXPfEx5Bt6oQ4n%2B%2BvC8hRgeWRUtHJfb6Q3LQMjahNDJTC95NoeEglgDqs4GdtnV%2FVojuRBVT4yS%2BetWqXyIP9qHZpOANDAJ%2BJ0gyEa6by1ONrUzEniplb5sqOyDBS5R4PnvIhVtivnM%2FZi6c2ISqRy%2FGPOzLwe0oQ%2F3XDpuJNPxnrwuXNR%2BOdyNLZcHuv4HBur3XSu1KDSzomFJw9rsV4UrD7D8z4UshUNgaZEjdDJfVbR6Jo7f1klq7kPvcde%2BoSxZ3v3NNtl6vaslalWY9k5wH69cnNCiZDRloJsxl5a%2BCrCAq49yXCtku0cZOWQfTbG28A0%2F2wvpXoAHvDJBGdxTB%2BO9WfYFXapIsi5vtzUWf%2Fan2Uv%2F9C96z14gnHnw1uu5eS5xSR9hY2txiPOt4XD6pxYqj4Sdds3ObhWfK%2BhTdrBFKKZmWA4g0ORDZhmAq6JeVmVVuCXrO9TJu16fTiM9IKh%2BdSfy7vDWTy3rO8BuXcyvK3LGv%2FpZ2nKFxm9R6VVBBkBm39W9TAwsoYI9ilC4XxSJ91%2FhKF5jurAmbj3%2F7lK9ukNPFHdhjvbsPBO999egKed34augmIkIa7xQkhWvAKPzGcZYbrx5B1QVxl25%2FKXPikYazPZwZX2F%2FY7FAsyLvUf7nYDoYGWg4XzUL4MD1Ypdgmn8aNgu6xsWDmwUCm2Ox3iVXrdd3APoEuLmlafsXvN36FGSa99FROQuadPUCUI0go6EA9tnxvGgJTVLkLVdP7C2ij0QvO9GXfULVlSUU5i0qZzeKAb%2FSGHLIN4xUp5qgPLnCNX7i5sXUA%2B%2Fig1%2BPpNZq8oh5kfRmNxZDc6jdGjpPXVEvVgkFGEX%2FudWv3Q%2BdQO17QYA26veJuy7YQ%2Ba0tJEUCo6Z07uvsqqeDI0pR8Mr24IJ1Ouf%2FgXghUTJnLwxe%2F0cJm0GuXYSTXpsj%2Bx4lucdmwxrthrCS78vwI3lMVylXNFjJzcrnvthYp39tW4wQx0KyBFsvCuNdg0mSc67SBIsnfkN5U%2BlAzUPmxmr7URLaZp%2BRtxkLM5n76uXG8kVZFTkpmyjRFvjzbEihUiLTLFTQBBiUUISOrskfhgea61DdJQwrA6seNL9XBbj2RGW5iCTxB7WWiJkr9SYr4lP3abRv2e8LIT8EcRh6cNVExmOurMxk1l6EDoTXZJ6RFS%2FjJgktWyXoww%2Fzpm9MbJvYvKQmmyTWIkJj3WSMUnI7cDrisjQ9MIKPCUjjTU37pYjqbEht6mpW45ueiuovDBWlTYYHqkVUcY3tLeV5BswEcblrHYULaw8%2F%2BKxyyKH19bPTGNVfr%2B%2BhsOogruLOxEGs2Zbndv1JCubZW%2FYz2Q%2Bl27AyrYUFZ1bjDj1bTA%2Bzil9KAO%2B58GcX5Dvvgi4LLk%2FBxTWVqr1XVlgwChx4ROoiV9yqeKwUXH%2BHLBrSxl7c0VbPQpo387JMHb0T64OhHbO6Qejr%2FmX%2BadbqGo3POAsyfW3LNA12aUjnlix%2FwpXcOAElrPXJu6J3zaP2jG0v8fXTo7jYhbCWexWCIsYRbh8aQJeESd3Y%2FdzwJ8wK3SPa2gpjD9xt%2FUSaCe2tw2nkuQfDiDr1WNiJLHV4noAIbY4%2BpnJTKSAK%2BQSeoyxT3AgieZeip0WC4hRRoAFn9k7kE3fvloa2n8%2B0zvmGgBxgzCbI83uoxtdt3ChvY4kySK78irkiE5d825S5nsKP3wkHZ6rlcnWHt2%2FVvB7yjC751OcUCbgAXfovbVmzatsAmwu5aVdH3W%2F4rYf3o%2BPXYuFhZD8BCACzeo0GaCtp9CzsZmZ%2F4chJAq7Qb0%2FHcRgogqPsBjBkkT%2BDbQKron3HVySJ0kIAkDJkynia9182OVlPz8Z9FwjAw2f1MMT%2Bd94qa4pKti5yqsFM4ekGgaXNc9EJ8Y2HPMjV8Lt7b6qdcMLC0r5hDCKZYErD3hi74ankge9uI07cRAEIMEbXplHkvYmzRReXJ5qrkfoUXGzeX421w2YhzaCX1F82KIh7vIJ2C3Cg9u5tWsOpApbS8LPPY64sf8seN%2FrHmBZxtONkYYHLhPN%2B4yupBkWSmXqwyNth%2Bj4JKmVamcWLFrgdvTWwSZistqceLCoQdAppAEtljXFHRQy1QoUNsOCwCs5CQ5SVOzaoPs1%2BqSR9LGy3pA37lbnyWmAYJ%2FWjmroi9MCB1b2EUQO70kevkDISt6ZeMSSPkT9Kbewspuxuwknv8RYUVHMtxJP5ss6pUGwtGsa4D6XrfYq4HPIoipcJNTS76Uzow7aphBuUvUbWCt6xV1xoH%2BqwrUMSOSAJ6f0fLWtLvoQfvcaHyNglCvZKYWcXsyQZL5GLCbrrng2eQEhf3Ld71TwSwm34mKPP2U9rvMRX9ldnb%2BySboz5RY9M15ZA6z0OpaTKH1hUhZqGwWm1OY8kiNEnyDVU8G7%2FRiAWTYK6NsW908t6AgqMCQ%3D";
        String a2 = "7D401E16";
        String a3 = "3X3YVyGwyKSG%2BxN9stB7c0yeyS7gEU%2Fjt0cNCIRQ0cMOGdqyvjxWR8H7riN%2FV8TrcG8qaToYNbkRFD5ZkV0H4SsHFobsEhkuwt8YZUrM%2BMXGN3nSUxisBUDfqqYcsTM2%2FssDMkFMrAtXUKJPmTvMVKgn8XJyCXSAvfKlFmH%2FX28fAP41%2FjVlw6TBQVhujpjPIOIn4ztW7ZGlNBtFd0JXBpOlOD%2FWB8IA8Y1lGp%2FiiNuFPKMjS4QSrQ%2BGjSAeqEbL8uugq5KMpEEsNAxnUK4DA69SngW5aghCeOHaX3UQk8TPhd%2BY18j2aI9wGoOLGCaoEXBUjR%2Bpl2KVOOyJZ%2BrTspEsPZyqnuFDWcK743ti3xkX%2FmLyvm2GAB4JqJLxZLSQXNQg7xxtW65QtyINv%2BdIT8MjRkJmhXUf9q2T9S0zN9CcODnms6mANa2e4ra9GAVrRKAKnsKYgphlzij6BjgQjlTfGSl0Os10vUceR6n%2FtRRRTDc2BaA01bczY2pw8Dfvcw6%2BRFXLW0TSCToEge26GUulNUSFHkoiESoNpA2HS2iZRJNCaKzBeOQ%2B%2Fo%2BeWNWk28AM5sFiLmc21ghCcbrjik%2BGm%2BMDVl2%2ByPL5TciBBH97JRQrSZvCG3zAnzNjobj2UfKpVEpW2naQe6FZF7m5Wo5%2FHCE6oLgm6mNR0Kv4bdIeRLqEWJDEgFspXP8%3D";

//231

        String url = "http://xgb.lzu.edu.cn/SystemForm/OnLine/VoteReport.aspx?Id="+id+"&Form=OnLineTopic.aspx";

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

        private static void GetInf(int ii,int jj,String code){

            Document doc = Jsoup.parse(code, "UTF-8");


//            String cookie2 = doc.getElementsByClass("aspNetHidden").get(0).select("input").get(2).attr("value");
//            String cookie1 = doc.getElementsByClass("aspNetHidden").get(1).select("input").get(1).attr("value");
//            String cookie3 = doc.getElementsByClass("aspNetHidden").get(1).select("input").get(0).attr("value");
//            System.out.println(cookie1+cookie2+cookie3);
//
            String A = "";
            String B = "";
            String C = "";
            String D = "";
            String E = "";
            String F = "";
            String aa [] = {A,B,C,D,E,F};

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
                            aa[j] = aa[j] + QItemPeopleNum + ",";

                        } catch (Exception e) {

                            e.printStackTrace();

                        }

//                    System.out.println(QItem);

//                    System.out.println(QItem.split(" ")[0] + "  " + QItemPeopleNum);
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
                System.out.println(CollegeAll[ii]+GradeAll[jj]+"========解析错误：");

            }

            try {
                CreateExcel(CollegeAll[ii], jj, aa, ALLPeopleNum);
            } catch (IOException | WriteException e) {
                e.printStackTrace();
            }

        }

        private static String getMatcher(String regex, String source) {
            String result = "";
            Pattern pattern = Pattern.compile(regex);
            Matcher matcher = pattern.matcher(source);
            while (matcher.find()) {
                result = matcher.group(1);
            }
            return result;
        }


        private static void CreateExcel(String file_name,int Sheet,String[] QPeopleNum,String ALLPeopleNum)
                throws IOException, WriteException {

            String Question[] = {"A","B","C","D","E","F"};
            //1:创建excel文件

            String path = "D:/"+file_name+".xls";

            File file=new File(path);
            WritableWorkbook workbook;
            if (!file.exists()) {
                file.createNewFile();
                workbook = Workbook.createWorkbook(file);

            }else {

                FileInputStream in = new FileInputStream(path);
                Workbook wb = null;
                try {
                    wb = Workbook.getWorkbook(in);
                } catch (BiffException e) {
                    e.printStackTrace();
                }
                workbook = wb.createWorkbook(new File(path),wb);
                //sheetname2 ，新建sheet的名称。

            }

            //2:创建工作簿

            //3:创建sheet,设置第二三四..个sheet，依次类推即可
            WritableSheet sheet = workbook.createSheet(GradeAll[Sheet]+ALLPeopleNum+"人", Sheet);
            //4：设置titles
            //5:单元格
            Label label=null;
            //6:给第一列设置列名
            for(int i=0; i<Q_titles.length; i++){
                //x,y,第一行的列名
                label=new Label(0, i, Q_titles[i]);
                //7：添加单元格
                sheet.addCell(label);
            }
            //导入数据
            for(int i = 0; i < QPeopleNum.length; i++){
                //添加编号，
                label=new Label(i+1, 0, Question[i]);
                sheet.addCell(label);
                System.out.println(QPeopleNum[i]);

                for (int j = 0; j < QPeopleNum[i].split(",").length; j++){
                    //添加账号
                    label=new Label(i+1, j+1, QPeopleNum[i].split(",")[j]);
                    sheet.addCell(label);

                }

            }

            //写入数据，一定记得写入数据，不然你都开始怀疑世界了，excel里面啥都没有
            workbook.write();
            //最后一步，关闭工作簿

            workbook.close();
        }


}
