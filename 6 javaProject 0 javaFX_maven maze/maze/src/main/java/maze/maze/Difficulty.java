package maze.maze;

public class Difficulty {

	private int size;
	private int startDelay;
	private int stepDelay;
	
	public Difficulty(int difficulty) {
		
        switch (difficulty) {
            case (1):
                size = 10;
            	startDelay = 3000;
            	stepDelay = 300;
                break;
            case (2):
                size = 15;
	            startDelay = 5000;
	            stepDelay = 300;
                break;
            case (3):
                size = 20;
            	startDelay = 10000;
            	stepDelay = 300;
                break;
            case (4):
                size = 25;
            	startDelay = 20000;
            	stepDelay = 300;
                break;
            case (5):
                size = 30;
            	startDelay = 40000;
            	stepDelay = 300;
                break;
            case (6):
                size = 35;
            	startDelay = 80000;
            	stepDelay = 300;
                break;
            case (7):
                size = 30;
            	startDelay = 10000000;
            	stepDelay = 10000000;
                break;
            case (8):
                size = 40;
            	startDelay = 10000000;
            	stepDelay = 10000000;
                break;
            case (9):
                size = 50;
            	startDelay = 10000000;
            	stepDelay = 10000000;
                break;
            default:
                System.out.println("error");
                break;
        }
        
        System.out.println("Difficulty created");
        
	}

	public int getSize() {
		return size;
	}
	public int getStartDelay() {
		return startDelay;
	}
	public int getStepDelay() {
		return stepDelay;
	}
	
}
