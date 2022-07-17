package maze.maze;

import org.jgrapht.graph.DefaultEdge;
import org.jgrapht.graph.SimpleGraph;

public class Static {
	
	// mod i
	
	public static void leftButton(Player player) {
		player.setPlayerWorldSide(player.getPlayerWorldSide()-1);
		worldSideCorrection(player);
		
		System.out.println("Static, left button");
		
	}
	public static void rightButton(Player player) {
		player.setPlayerWorldSide(player.getPlayerWorldSide()+1);
		worldSideCorrection(player);
		
		System.out.println("Static, right button");
		
	}
	public static void worldSideCorrection(Player player) {
		if(player.getPlayerWorldSide()<1) {
			player.setPlayerWorldSide(player.getPlayerWorldSide()+4);
		}
		if(player.getPlayerWorldSide()>4) {
			player.setPlayerWorldSide(player.getPlayerWorldSide()-4);
		}
	}
	
	public static void buttonFront(Maze maze, Player player) {
        switch (player.getPlayerWorldSide()) {
            case (1): // east
            	boolean[][] east = maze.getEast();
	            if(east[player.getPlayerPositionX()][player.getPlayerPositionY()]==true) {
	            	break;
	            }
	            player.setPlayerPositionX(player.getPlayerPositionX()+1);
                break;
            case (2): // south
            	boolean[][] south = maze.getSouth();
	            if(south[player.getPlayerPositionX()][player.getPlayerPositionY()]==true) {
	            	break;
	            }
	            player.setPlayerPositionY(player.getPlayerPositionY()-1);
                break;
            case (3): // west
            	boolean[][] west = maze.getWest();
	            if(west[player.getPlayerPositionX()][player.getPlayerPositionY()]==true) {
	            	break;
	            }
	            player.setPlayerPositionX(player.getPlayerPositionX()-1);
                break;
            case (4): // north
            	boolean[][] north = maze.getNorth();
	            if(north[player.getPlayerPositionX()][player.getPlayerPositionY()]==true) {
	            	break;
	            }
	            player.setPlayerPositionY(player.getPlayerPositionY()+1);
                break;
            default:
                System.out.println("greska");
                break;
        }
        
        System.out.println("Static, front button");
        
	}
	
	public static void buttonBack(Maze maze, Player player) {
        switch (player.getPlayerWorldSide()) {
            case (1): // east
            	boolean[][] west = maze.getWest();
	            if(west[player.getPlayerPositionX()][player.getPlayerPositionY()]==true) {
	            	break;
	            }
	            player.setPlayerPositionX(player.getPlayerPositionX()-1);
                break;
            case (2): // south
            	boolean[][] north = maze.getNorth();
	            if(north[player.getPlayerPositionX()][player.getPlayerPositionY()]==true) {
	            	break;
	            }
	            player.setPlayerPositionY(player.getPlayerPositionY()+1);
                break;
            case (3): // west
            	boolean[][] east = maze.getEast();
	            if(east[player.getPlayerPositionX()][player.getPlayerPositionY()]==true) {
	            	break;
	            }
	            player.setPlayerPositionX(player.getPlayerPositionX()+1);
                break;
            case (4): // north
            	boolean[][] south = maze.getSouth();
	            if(south[player.getPlayerPositionX()][player.getPlayerPositionY()]==true) {
	            	break;
	            }
	            player.setPlayerPositionY(player.getPlayerPositionY()-1);
                break;
            default:
                System.out.println("greska");
                break;
        }
        
        System.out.println("Static, back button");
        
	}
	
	// marker
	
	public static void setMarker(Convert jgraphtCreate, SimpleGraph<GNode, DefaultEdge> convertedGraph, Player player) {
		GNode extractedNode = jgraphtCreate.getJgraphtNode(player.getPlayerPositionX(), player.getPlayerPositionY(), convertedGraph);
		extractedNode.setMarker(1);

		System.out.println("Static, marker button");
		
	}
	
	// mod ii
	
	// north
	public static void upArrow(Maze maze, Player player) {
		boolean[][] north = maze.getNorth();
		player.setPlayerPreviousX(player.getPlayerPositionX());
		player.setPlayerPreviousY(player.getPlayerPositionY());
		if(north[player.getPlayerPositionX()][player.getPlayerPositionY()]==false) {
			player.setPlayerPositionY(player.getPlayerPositionY()+1);
        }
		
		System.out.println("Static, arrow up");
		
	}
	
	// south
	public static void downArrow(Maze maze, Player player) {
		boolean[][] south = maze.getSouth();
		player.setPlayerPreviousX(player.getPlayerPositionX());
		player.setPlayerPreviousY(player.getPlayerPositionY());
		if(south[player.getPlayerPositionX()][player.getPlayerPositionY()]==false) {
			player.setPlayerPositionY(player.getPlayerPositionY()-1);
        }
		
		System.out.println("Static, arrow down");
		
	}
	
	// east
	public static void rightArrow(Maze maze, Player player) {
		boolean[][] east = maze.getEast();
		player.setPlayerPreviousX(player.getPlayerPositionX());
		player.setPlayerPreviousY(player.getPlayerPositionY());
		if(east[player.getPlayerPositionX()][player.getPlayerPositionY()]==false) {
			player.setPlayerPositionX(player.getPlayerPositionX()+1);
        }
		
		System.out.println("Static, arrow right");
		
	}
	
	// west
	public static void leftArrow(Maze maze, Player player) {
		boolean[][] west = maze.getWest();
		player.setPlayerPreviousX(player.getPlayerPositionX());
		player.setPlayerPreviousY(player.getPlayerPositionY());
		if(west[player.getPlayerPositionX()][player.getPlayerPositionY()]==false) {
			player.setPlayerPositionX(player.getPlayerPositionX()-1);
        }
		
		System.out.println("Static, arrow left");
		
	}

}
