import java.util.Scanner;
public class ArrayUTILITY {
    public static int[] inputArray() {
        Scanner input = new Scanner(System.in);
        System.out.println("Enter 5 integers:");
        int size = input.nextInt();
        int arr[] = new int[size];
        int i = 0;
        while(i< size) {
            System.out.println("Enter element no." + (i+1) + ": ");
            arr[i] = input.nextInt();
            i++;
        }
        return arr;
    }
}