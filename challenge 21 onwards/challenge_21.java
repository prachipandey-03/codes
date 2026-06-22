import java.util.Scanner;
public class challenge_21 {
    public static void main(String[] args) {
        Scanner input = new Scanner (System.in);
        System.out.println("Welcome to bitwise operator checker");
        System.out.println("Enter the first number");
        int num1 = input.nextInt();
        System.out.println("Enter the second number");
        int num2 = input.nextInt();
       // System.out.println("Bitwise AND: " + (num1 & num2));
        System.out.println("Bitwise XOR: " + (num1 ^ num2));
        int result = num1 & num2;
    }
        //int result = num1 & num2;
       // System.out.println("Bitwise OR: " + (num1 | num2));
       // System.out.println("Bitwise XOR: " + (num1 ^ num2));

    }
    


