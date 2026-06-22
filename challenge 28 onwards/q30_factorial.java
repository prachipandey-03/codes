import java.util.Scanner;
public class q30_factorial {
    public static void main(String[] args) {
        Scanner input = new Scanner(System.in);
        System.out.println("Factorial of a number");
        System.out.println("Enter a number N: ");
        int num = input.nextInt();
        long result = factorial(num);
        System.out.println("Factorial of " + num + " is: " + result);

    }
    public static long factorial(int num) {
        long fact = 1;
        for(int i = 1; i <= num; i++) {
            fact *= i;
        }
        return fact;
    }
}
      