import java.util.Scanner;
public class challenge_16 {
    public static void main(String[] args) {
        Scanner input = new Scanner(System.in);
        System.out.println("Welcome to the Even/Odd checker!");
        System.out.println("Enter a number:");
        int num = input.nextInt();

        if (num%2 ==0) {
            System.out.println("The number is even.");
            
    }   else {
        System.out.println("The number is odd.");
    }
}

    
}