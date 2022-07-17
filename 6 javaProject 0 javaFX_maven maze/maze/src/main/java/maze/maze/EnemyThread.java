package maze.maze;

import java.util.Iterator;
import java.util.List;

import org.jgrapht.alg.shortestpath.DijkstraShortestPath;
import org.jgrapht.graph.DefaultEdge;
import org.jgrapht.graph.SimpleGraph;

public class EnemyThread implements Runnable {
	
	private Player player;
	private Enemy enemy;
	private Difficulty difficulty;
	private Convert jgraphtCreate;
	private SimpleGraph<GNode, DefaultEdge> convertedGraph;
	private InterfaceData interfaceData;

	public EnemyThread(Player player, Enemy enemy, Difficulty difficulty, Convert jgraphtCreate,
			SimpleGraph<GNode, DefaultEdge> convertedGraph, InterfaceData interfaceData) {
		this.player = player;
		this.enemy = enemy;
		this.difficulty = difficulty;
		this.jgraphtCreate = jgraphtCreate;
		this.convertedGraph = convertedGraph;
		this.interfaceData = interfaceData;

		System.out.println("EnemyThread created");
		
	}

	@Override
	public void run() {
		
		// wait at start
		try {
			Thread.sleep(difficulty.getStartDelay());
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
		
		while(true) {
			
			if(interfaceData.getGameLiveOrNot() == 0) {
	        	break;
	        }
			
			// wait for step
			try {
				Thread.sleep(difficulty.getStepDelay());
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
			
			// get player position in jgrapht
			GNode jgraphtNodePlayer = jgraphtCreate.getJgraphtNode(player.getPlayerPositionX(), player.getPlayerPositionY(), convertedGraph);
			
			System.out.println("enemyThread, player node " + jgraphtNodePlayer.getPosX() + " " + jgraphtNodePlayer.getPosY());
			
			// get enemy position
			GNode jgraphtNodeEnemy = jgraphtCreate.getJgraphtNode(enemy.getEnemyPositionX(), enemy.getEnemyPositionY(), convertedGraph);
			
			System.out.println("enemyThread, enemy node " + jgraphtNodeEnemy.getPosX() + " " + jgraphtNodeEnemy.getPosY());

			// if catch player end the game and exit while loop
	        if(player.getPlayerPositionX()==enemy.getEnemyPositionX() && player.getPlayerPositionY()==enemy.getEnemyPositionY()) {
	        	player.setPlayerLife(0);
	        	break;
	        }
			
			// calculate shortest path and return range of steps
			DijkstraShortestPath<GNode, DefaultEdge> dijkstraShortestPath = new DijkstraShortestPath<>(convertedGraph);
	        List<GNode> shortestPathVertices = dijkstraShortestPath.getPath(jgraphtNodeEnemy,jgraphtNodePlayer).getVertexList();
	        
	        System.out.println("enemyThread, dijkstra " + dijkstraShortestPath.toString() + " " + shortestPathVertices.toString());
	        
	        // return one closer step coordinates
	        Iterator<GNode> iter = shortestPathVertices.iterator();
	        iter.next();
	        GNode oneCloser = (GNode) iter.next();

	        System.out.println("enemyThread, one closer node " + oneCloser.toString());
	        
	        // save enemy previous position for drawing
	        enemy.setEnemyPreviousX(enemy.getEnemyPositionX());
	        enemy.setEnemyPreviousY(enemy.getEnemyPositionY());
	        
	        // move enemy to calculated position
	        enemy.setEnemyPositionX(oneCloser.getPosX());
	        enemy.setEnemyPositionY(oneCloser.getPosY());

		}
		
		System.out.println("enemyThread, game ends");
		
	}

}
