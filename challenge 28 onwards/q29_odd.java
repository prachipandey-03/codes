import java.util.Scanner;
public class q29_odd {
    public static void main(String[] args) {
        Scanner input = new Scanner(System.in);
        System.out.println("Sum of odd numbers");
        System.out.println("Enter a number N: ");
        int num = input.nextInt();
        int result = sumofodd(num);
        System.out.println("Sum of odd numbers from 1 to " + num + " is: " + result);

    }
    public static int sumofodd(int num) {
        int sum = 0;
        int i = 1;
        while(i<=num) {
            
            sum+=i;
            i+=2;
            // Move to the next odd number
        }
        return sum ;
    }
}
