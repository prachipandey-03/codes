import java.util.Scanner;
public class challenge_15 {
    public static void main(String[] args) {
        Scanner input = new Scanner(System.in);
        System.out.print("welcome to Number Checker\n ");
        System.out.println("Enter a number: ");
        int num = input.nextInt();

        if (num > 0) {
            System.out.println("The number is positive.");
        } else if (num == 0){
            System.out.println("The number is zero.");
        } else {
            System.out.println("The number is negative.");
        }
    }
}
