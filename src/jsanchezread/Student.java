package jsanchezread;

import java.util.Objects;

/**
 * Created by anirudh on 20/10/14.
 */
public class Student {

    private String name;
    private String maths;
    private String science;
    private String english;

    public Student(){}

    public Student(String name, String maths, String science, String english) {
        this.name = name;
        this.maths = maths;
        this.science = science;
        this.english = english;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getMaths() {
        return maths;
    }

    public void setMaths(String maths) {
        this.maths = maths;
    }

    public String getScience() {
        return science;
    }

    public void setScience(String science) {
        this.science = science;
    }

    public String getEnglish() {
        return english;
    }

    public void setEnglish(String english) {
        this.english = english;
    }

    @Override
    public int hashCode() {
        int hash = 7;
        hash = 97 * hash + Objects.hashCode(this.name);
        return hash;
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj) {
            return true;
        }
        if (obj == null) {
            return false;
        }
        if (getClass() != obj.getClass()) {
            return false;
        }
        final Student other = (Student) obj;
        if (!Objects.equals(this.name, other.name)) {
            return false;
        }
        return true;
    }

    
        
    @Override
    public String toString() {
               return name+ ": Maths "+maths+ " Science "+science+" English "+english;

    }
}
