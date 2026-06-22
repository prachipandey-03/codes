import java.lang.*;
public class typeconversion {
    public static void main(String[] args) {
        float myFloat = 5;  //converting lower  to higher type (implicit)
        System.out.println(myFloat); 
        int myInt = (int) 5.0f; //converting higher to lower type (explicit)
    }    
}
