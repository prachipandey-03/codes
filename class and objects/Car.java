public class Car {
    public static void main(String[] args) {
        int noOfWheels;
        String color;
        float maxSpeed;
        int noOfSeats;
        
        public void drive() {
                System.out.println("The car is driving.")
                currentfuelinlitres--;


        }
        public void addFuel(float fuel) {
            currentfuelinlitres += fuel;
        } 
        public float getCurrentFuel() {
            return currentFuelinlitres;


        }
    }
}