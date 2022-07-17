package maze.maze;

public class Enemy {

	private int enemyPositionX;
	private int enemyPositionY;
	private int enemyPreviousX;
	private int enemyPreviousY;

	public Enemy() {
		System.out.println("Enemy created");
	}
	
	public int getEnemyPositionX() {
		return enemyPositionX;
	}
	public void setEnemyPositionX(int enemyPositionX) {
		this.enemyPositionX = enemyPositionX;
	}
	public int getEnemyPositionY() {
		return enemyPositionY;
	}
	public void setEnemyPositionY(int enemyPositionY) {
		this.enemyPositionY = enemyPositionY;
	}
	public int getEnemyPreviousX() {
		return enemyPreviousX;
	}
	public void setEnemyPreviousX(int enemyPreviousX) {
		this.enemyPreviousX = enemyPreviousX;
	}
	public int getEnemyPreviousY() {
		return enemyPreviousY;
	}
	public void setEnemyPreviousY(int enemyPreviousY) {
		this.enemyPreviousY = enemyPreviousY;
	}

}
