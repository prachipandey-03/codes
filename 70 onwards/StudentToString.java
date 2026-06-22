public class StudentToString {
    
    public static void main(String[] args) {
        String name;
        int age;
        String rollNumber;
        String house;

        public StudentToString(String name, int age, String rollNumber, String house) { 
            this.name = name;
            this.age = age;
            this.rollNumber = rollNumber;
            this.house = house;
        }

        @Override
        public String toString() {
            return "Student Details:{name: , age: , rollNumber: , house: }" 
        }


    }
     
}
