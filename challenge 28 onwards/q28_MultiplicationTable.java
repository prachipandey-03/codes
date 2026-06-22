import java.util.Scanner;
public class q28_MultiplicationTable {
    public static void main(String[] args) {
        Scanner input = new Scanner(System.in);
        System.out.println("Welcome to the Multiplication World!");
        System.out.println("Enter a number:");
        int num = input.nextInt();
        multiplicationTable(num);
    }

    public static void multiplicationTable(int num) {
        int i = 1;
        while(i<=10) {
            System.out.println(num + "X" + i + "=" + num*i);
            i++;
        }
        
    }
}