
package maze.maze;

import org.jgrapht.graph.DefaultEdge;
import org.jgrapht.graph.SimpleGraph;
import org.jgrapht.traverse.BreadthFirstIterator;

import javafx.geometry.Rectangle2D;
import javafx.scene.canvas.GraphicsContext;
import javafx.scene.paint.Color;
import javafx.stage.Screen;

public class Game {

	GraphicsContext gc;
	InterfaceData interfaceData;
	Difficulty difficulty;
	Maze maze;
	Convert jgraphtConvert;
	SimpleGraph<GNode, DefaultEdge> convertedGraph;
	GNode[][] helpMatrix;
	GNode exitNode;
	Player player;
	Enemy enemy;

	Player player2;
	Enemy enemy2;

	public GraphicsContext getGc() {
		return gc;
	}
	public InterfaceData getInterfaceData() {
		return interfaceData;
	}
	public Difficulty getDifficulty() {
		return difficulty;
	}
	public void setDifficulty(Difficulty difficulty) {
		this.difficulty = difficulty;
	}
	public Maze getMaze() {
		return maze;
	}
	public void setMaze(Maze maze) {
		this.maze = maze;
	}
	public Convert getJgraphtConvert() {
		return jgraphtConvert;
	}
	public void setJgraphtConvert(Convert jgraphtConvert) {
		this.jgraphtConvert = jgraphtConvert;
	}
	public SimpleGraph<GNode, DefaultEdge> getConvertedGraph() {
		return convertedGraph;
	}
	public void setConvertedGraph(SimpleGraph<GNode, DefaultEdge> convertedGraph) {
		this.convertedGraph = convertedGraph;
	}
	public GNode[][] getHelpMatrix() {
		return helpMatrix;
	}
	public void setHelpMatrix(GNode[][] helpMatrix) {
		this.helpMatrix = helpMatrix;
	}
	public GNode getExitNode() {
		return exitNode;
	}
	public void setExitNode(GNode gNode) {
		this.exitNode = gNode;
	}
	public Player getPlayer() {
		return player;
	}
	public void setPlayer(Player player) {
		this.player = player;
	}
	public Enemy getEnemy() {
		return enemy;
	}
	public void setEnemy(Enemy enemy) {
		this.enemy = enemy;
	}
	public Player getPlayer2() {
		return player2;
	}
	public void setPlayer2(Player player2) {
		this.player2 = player2;
	}
	public Enemy getEnemy2() {
		return enemy2;
	}
	public void setEnemy2(Enemy enemy2) {
		this.enemy2 = enemy2;
	}
	
	public Game(GraphicsContext gc, InterfaceData interfaceData) {
		this.gc = gc;
		this.interfaceData = interfaceData;
		
		System.out.println("game created");
		
	}
	
	public void playGame() {

		difficulty = new Difficulty(interfaceData.getInputedDifficulty());
    	int n = difficulty.getSize();

        maze = new Maze(n);

        jgraphtConvert = new Convert(maze);
    	convertedGraph = jgraphtConvert.getConvertedGraph();
		helpMatrix = jgraphtConvert.getHelpMatrix();
		
		exitNode = jgraphtConvert.getJgraphtExitNode(1, convertedGraph);
		
		player = new Player();
		player.setPlayerPositionX(1);
		player.setPlayerPositionY(n);
		player.setPlayerWorldSide(1);
		player.setPlayerPreviousX(1);
		player.setPlayerPreviousY(n);
		player.setPlayerLife(1);
		
		enemy = new Enemy();
		enemy.setEnemyPositionX(1);
		enemy.setEnemyPositionY(n);
		enemy.setEnemyPreviousX(1);
		enemy.setEnemyPreviousY(n);
		
		EnemyThread enemyRunnable = new EnemyThread(player, enemy, difficulty, jgraphtConvert, convertedGraph, interfaceData);
        Thread nit = new Thread(enemyRunnable);
        nit.start();
		
        if (interfaceData.getNumberOfPlayers() == 2) {
        	
        	player2 = new Player();
    		player2.setPlayerPositionX(1);
    		player2.setPlayerPositionY(n);
    		player2.setPlayerWorldSide(1);
    		player2.setPlayerPreviousX(1);
    		player2.setPlayerPreviousY(n);
    		player2.setPlayerLife(1);
    		
    		enemy2 = new Enemy();
    		enemy2.setEnemyPositionX(1);
    		enemy2.setEnemyPositionY(n);
    		enemy2.setEnemyPreviousX(1);
    		enemy2.setEnemyPreviousY(n);
    		
    		EnemyThread enemyRunnable2 = new EnemyThread(player2, enemy2, difficulty, jgraphtConvert, convertedGraph, interfaceData);
            Thread nit2 = new Thread(enemyRunnable2);
            nit2.start();
            	
        }

		System.out.println("Game created");
		
	}

	public void drawing() {
		
    	// unpack
    	int n = maze.getN();
    	boolean[][] north = maze.getNorth();
        boolean[][] east = maze.getEast();
        boolean[][] south = maze.getSouth();
        boolean[][] west = maze.getWest();

        // drawing setup
        gc.setLineWidth(2);
        
        // get screen resolution
    	Rectangle2D primaryScreenBounds = Screen.getPrimary().getVisualBounds();
        // recalculate for retrieved resolution
        double ct =  (primaryScreenBounds.getHeight()) / (difficulty.getSize() + 2);
        
        // walls
        gc.setFill(Color.BLUE);
        gc.setStroke(Color.BLUE);
        for (int x = 1; x <= n; x++) {
            for (int y = 1; y <= n; y++) {
                if (south[x][y]) gc.strokeLine(x*ct, y*ct, (x+1)*ct, y*ct);
                if (north[x][y]) gc.strokeLine(x*ct, (y+1)*ct, (x+1)*ct, (y+1)*ct);
                if (west[x][y])  gc.strokeLine(x*ct, y*ct, x*ct, (y+1)*ct);
                if (east[x][y])  gc.strokeLine((x+1)*ct, y*ct, (x+1)*ct, (y+1)*ct);
            }
        }
        
        // start & exit & marker
        BreadthFirstIterator<GNode, DefaultEdge> breadthFirstSearchIterator = new BreadthFirstIterator<>(convertedGraph);
		while(breadthFirstSearchIterator.hasNext()) {
        	GNode jgraphtNode = breadthFirstSearchIterator.next();
        	if(jgraphtNode.getStart()==1 || jgraphtNode.getExit()==1 || jgraphtNode.getMarker()==1) {
        		if(jgraphtNode.getStart()==1) {
            		gc.setFill(Color.GRAY);
            	}
            	if(jgraphtNode.getExit()==1) {
            		gc.setFill(Color.GREEN);
            	}
            	if(jgraphtNode.getMarker()==1) {
            		gc.setFill(Color.BROWN);
            	}
        		gc.fillOval((jgraphtNode.getPosX() + 0.25)*ct, (jgraphtNode.getPosY() + 0.25)*ct, 0.5*ct, 0.5*ct);
        	}
		}
        
		// testing
//    	gc.setFill(Color.GOLD);
//		BreadthFirstIterator<GNode, DefaultEdge> breadthFirstSearchIterator5 = new BreadthFirstIterator<>(convertedGraph);
//		while(breadthFirstSearchIterator5.hasNext()) {
//        	GNode jgraphtNode = breadthFirstSearchIterator5.next();
//    		gc.fillOval((jgraphtNode.getPosX() + 0.25)*ct, (jgraphtNode.getPosY() + 0.25)*ct, 0.5*ct, 0.5*ct);
//		}
		
		// clear previous enemy & player position
 		gc.clearRect((player.getPlayerPreviousX() + 0.25)*ct, (player.getPlayerPreviousY() + 0.25)*ct, 0.5*ct, 0.5*ct);
 		gc.clearRect((enemy.getEnemyPreviousX() + 0.25)*ct, (enemy.getEnemyPreviousY() + 0.25)*ct, 0.5*ct, 0.5*ct);
 		
 		if(interfaceData.getNumberOfPlayers() == 2) {
 			gc.clearRect((player2.getPlayerPreviousX() + 0.25)*ct, (player2.getPlayerPreviousY() + 0.25)*ct, 0.5*ct, 0.5*ct);
 	 		gc.clearRect((enemy2.getEnemyPreviousX() + 0.25)*ct, (enemy2.getEnemyPreviousY() + 0.25)*ct, 0.5*ct, 0.5*ct);
 		}
 		
        // enemy & player position
        gc.setFill(Color.RED);
        gc.fillOval((enemy.getEnemyPositionX() + 0.25)*ct, (enemy.getEnemyPositionY() + 0.25)*ct, 0.5*ct, 0.5*ct);
        
        if(player.getPlayerLife() == 0) {
        	gc.setFill(Color.BLACK);
            gc.fillOval((player.getPlayerPositionX() + 0.25)*ct, (player.getPlayerPositionY() + 0.25)*ct, 0.5*ct, 0.5*ct);
        }else if(player.getPlayerLife() == 1) {
        	gc.setFill(Color.ORANGE);
            gc.fillOval((player.getPlayerPositionX() + 0.25)*ct, (player.getPlayerPositionY() + 0.25)*ct, 0.5*ct, 0.5*ct);
        }
        
        if(interfaceData.getNumberOfPlayers() == 2) {
        	gc.setFill(Color.RED);
        	gc.fillOval((enemy2.getEnemyPositionX() + 0.25)*ct, (enemy2.getEnemyPositionY() + 0.25)*ct, 0.5*ct, 0.5*ct);
        	
        	if(player2.getPlayerLife() == 0) {
        		gc.setFill(Color.BLACK);
            	gc.fillOval((player2.getPlayerPositionX() + 0.25)*ct, (player2.getPlayerPositionY() + 0.25)*ct, 0.5*ct, 0.5*ct);
        	}else if(player2.getPlayerLife() == 1) {
        		gc.setFill(Color.DARKORANGE);
            	gc.fillOval((player2.getPlayerPositionX() + 0.25)*ct, (player2.getPlayerPositionY() + 0.25)*ct, 0.5*ct, 0.5*ct);
        	}
		}
	}
}
