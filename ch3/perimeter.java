import java.util.Scanner;
public class perimeter {
    public static void main(String[] args) {
        Scanner input = new Scanner(System.in);
        System.out.println("Enter length1: ");
        int a = input.nextInt();
        System.out.println("Enter breadth1: ");
        int b = input.nextInt();
        System.out.println("Enter length2: ");
        int c = input.nextInt();
        System.out.println("Enter breadth2: ");
        int d = input.nextInt();
        int e = a+b+c+d;

        System.out.println("Perimeter of Rectangle is: "+e);
    }
}
