import java.util.Scanner;
public class q31_LCM{
    public static void main(String[] args) {
        Scanner input = new Scanner(System.in);
        System.out.println("LCM of a number");
        System.out.println("Enter a number N: ");
        int num = input.nextInt();
        System.out.println("Enter a number N: ");
        int num2 = input.nextInt();
        int result = LCM(num, num2);
        System.out.println("LCM of two numbers is: " + result);

    }
    public static int LCM(int num , int num2) {
            int i = 1;
            while(i <=num2) {
                int factor = num*i;
                if(factor%num2 == 0) {
                    return factor;
                }
                    i++; 
                    

                }
                return 0;
            }

        
    }

      