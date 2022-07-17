package maze.maze;

import org.jgrapht.graph.DefaultEdge;
import org.jgrapht.graph.SimpleGraph;
import org.jgrapht.traverse.BreadthFirstIterator;

public class Convert {
	
	Maze maze;
	
	SimpleGraph<GNode, DefaultEdge> convertedGraph;
	GNode[][] helpMatrix;

	public SimpleGraph<GNode, DefaultEdge> getConvertedGraph() {
		return convertedGraph;
	}
	public GNode[][] getHelpMatrix() {
		return helpMatrix;
	}

	public Convert(Maze maze) {
		
		this.maze = maze;
		
		int n = maze.getN();
		boolean[][] west = maze.getWest();
		boolean[][] south = maze.getSouth();

		convertedGraph = new SimpleGraph<>(DefaultEdge.class);
		helpMatrix = new GNode[n][n];
		
		// go through original maze (jgrapht)
		for(int i=1; i<n+1; i++) {
			for(int j=1; j<n+1; j++) {
				
				// create node & set graphNode id
				GNode gn = new GNode();
				int id = Integer.parseInt(Integer.toString(i)+Integer.toString(j));
				gn.setId(id);
				
				// set position of a node
				gn.setPosX(i);
				gn.setPosY(j);
				
				// set start & end
				if(i==1 && j==n) {
					gn.setStart(1);
				}
				if(i==n && j==1) {
					gn.setExit(1);
				}

				// add node to graph and help matrix (help matrix -1 because it's an array)
				convertedGraph.addVertex(gn);
				helpMatrix[i-1][j-1] = gn;

				// horizontal connection in graph (skip if first row)
				if (i!=1 && west[i][j]==false) {
					convertedGraph.addEdge(gn, helpMatrix[i-2][j-1]);
				}
	
				// vertical connection in graph (skip if first row)(norht switched with south ???????????)
				if (j!=1 && south[i][j]==false) {
					convertedGraph.addEdge(gn, helpMatrix[i-1][j-2]);
				}
			}
		}

		System.out.println("Convert created " + helpMatrix.toString());
		System.out.println("Convert created " + convertedGraph.toString());

	}
	
	// returns node for specified coordinates
	public GNode getJgraphtNode(int posX, int posY, SimpleGraph<GNode, DefaultEdge> convertedGraph) {
		BreadthFirstIterator<GNode, DefaultEdge> breadthFirstSearchIterator = new BreadthFirstIterator<>(convertedGraph);
		while(breadthFirstSearchIterator.hasNext()) {
        	GNode jgraphtNode = breadthFirstSearchIterator.next();
        	if(jgraphtNode.getPosX()==posX && jgraphtNode.getPosY()==posY) {
        		return jgraphtNode;
        	}
		}
		
		System.out.println("convert, node coordinates returned");
		
		return null;
	}
	
	// returns exit node
	public GNode getJgraphtExitNode(int exit, SimpleGraph<GNode, DefaultEdge> convertedGraph) {
		BreadthFirstIterator<GNode, DefaultEdge> breadthFirstSearchIterator = new BreadthFirstIterator<>(convertedGraph);
		while(breadthFirstSearchIterator.hasNext()) {
        	GNode jgraphtNode = breadthFirstSearchIterator.next();
        	if(jgraphtNode.getExit() == 1) {
        		return jgraphtNode;
        	}
		}
		
		System.out.println("convert, exit node returned");
		
		return null;
	}
	
}
