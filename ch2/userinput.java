import java.util.Scanner;

public class userinput {
    public static void main(String[] args) {
        Scanner input = new Scanner(System.in);
        System.out.print("Enter your name here:");
        String name = input.nextLine();
        System.out.println("Good Morning " + name); 
        System.out.println(name + " also tell me your age: ");
        int age = input.nextInt();
        System.out.println("your age is: " + age);
    }
    

}
