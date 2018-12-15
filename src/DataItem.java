import java.util.List;

public class DataItem {

    String College;
    String Gradle;
    String People;
    List<Que> queList;

    public String getCollege() {
        return College;
    }

    public void setCollege(String college) {
        College = college;
    }

    public String getGradle() {
        return Gradle;
    }

    public void setGradle(String gradle) {
        Gradle = gradle;
    }

    public String getPeople() {
        return People;
    }

    public void setPeople(String people) {
        People = people;
    }

    public List<Que> getQueList() {
        return queList;
    }

    public void setQueList(List<Que> queList) {
        this.queList = queList;
    }


}
