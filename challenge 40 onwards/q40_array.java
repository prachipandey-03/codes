import java.util.Scanner;
public class q40_array {
    public static void main(String[] args) {
        int numArray[] = ArrayUTILITY.inputArray();
        long sum = sum(numArray);
        int average = average(numArray);
        System.out.println("Sum of the array is: " + sum);
        System.out.println("Average of the array is: " + average);
    }
    public static long sum(int[] numArray) {
        long sum = 0;
        int i = 0;
        while(i < numArray.length) {
            sum += numArray[i];
            i++;
        }
        return sum;
    }
    public static int average(int[] numArray) {
        long sum = sum(numArray);
        return (int) (sum / numArray.length);
    }
}