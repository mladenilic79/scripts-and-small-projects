package maze.maze;

public class Player {

	private int playerPositionX;
	private int playerPositionY;
	private int playerWorldSide;
	private int playerLife;
	private int playerPreviousX;
	private int playerPreviousY;
	
	public Player() {
		System.out.println("Player created");
	}
	
	public int getPlayerPositionX() {
		return playerPositionX;
	}
	public void setPlayerPositionX(int playerPositionX) {
		this.playerPositionX = playerPositionX;
	}
	public int getPlayerPositionY() {
		return playerPositionY;
	}
	public void setPlayerPositionY(int playerPositionY) {
		this.playerPositionY = playerPositionY;
	}
	public int getPlayerWorldSide() {
		return playerWorldSide;
	}
	public void setPlayerWorldSide(int playarWorldSide) {
		this.playerWorldSide = playarWorldSide;
	}
	public int getPlayerPreviousX() {
		return playerPreviousX;
	}
	public void setPlayerPreviousX(int playerPreviousX) {
		this.playerPreviousX = playerPreviousX;
	}
	public int getPlayerPreviousY() {
		return playerPreviousY;
	}
	public void setPlayerPreviousY(int playerPreviousY) {
		this.playerPreviousY = playerPreviousY;
	}
	public int getPlayerLife() {
		return playerLife;
	}
	public void setPlayerLife(int playerLife) {
		this.playerLife = playerLife;
	}

}
