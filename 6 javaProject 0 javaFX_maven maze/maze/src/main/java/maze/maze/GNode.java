package maze.maze;

public class GNode {
	
	private int id;
	private int posX;
	private int posY;
	private int start=0;
	private int exit=0;
	private int marker=0;
	
	public GNode() {
	}
	
	public int getPosX() {
		return posX;
	}
	public void setPosX(int posX) {
		this.posX = posX;
	}
	public int getPosY() {
		return posY;
	}
	public void setPosY(int posY) {
		this.posY = posY;
	}
	public int getId() {
		return id;
	}
	public void setId(int id) {
		this.id = id;
	}
	public int getStart() {
		return start;
	}
	public void setStart(int start) {
		this.start = start;
	}
	public int getExit() {
		return exit;
	}
	public void setExit(int exit) {
		this.exit = exit;
	}
	public int getMarker() {
		return marker;
	}
	public void setMarker(int marker) {
		this.marker = marker;
	}

}
