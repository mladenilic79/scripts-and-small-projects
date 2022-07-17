package maze.maze;

public class InterfaceData {

	private int inputedDifficulty;
	private int mod;
	private int numberOfPlayers;
	private int gameLiveOrNot;
	
	public InterfaceData() {
		this.gameLiveOrNot = 0;
	}

	public int getInputedDifficulty() {
		return inputedDifficulty;
	}
	public void setInputedDifficulty(int inputedDifficulty) {
		this.inputedDifficulty = inputedDifficulty;
	}
	public int getMod() {
		return mod;
	}
	public void setMod(int mod) {
		this.mod = mod;
	}
	public int getNumberOfPlayers() {
		return numberOfPlayers;
	}
	public void setNumberOfPlayers(int numberOfPlayers) {
		this.numberOfPlayers = numberOfPlayers;
	}
	public int getGameLiveOrNot() {
		return gameLiveOrNot;
	}
	public void setGameLiveOrNot(int gameLiveOrNot) {
		this.gameLiveOrNot = gameLiveOrNot;
	}

}
