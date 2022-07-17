package maze.maze;

import javafx.animation.AnimationTimer;
import javafx.application.Application;
import javafx.application.Platform;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Pos;
import javafx.geometry.Rectangle2D;
import javafx.scene.Scene;
import javafx.scene.canvas.Canvas;
import javafx.scene.canvas.GraphicsContext;
import javafx.scene.control.Button;
import javafx.scene.control.ChoiceBox;
import javafx.scene.control.Label;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyEvent;
import javafx.scene.layout.GridPane;
import javafx.stage.Screen;
import javafx.stage.Stage;
import javafx.stage.WindowEvent;

public class App extends Application {

    @Override
    public void start(Stage stage) {
    	
		////////////////////////////////////////////////////////////////////////
		////////////////////////////////////////////////////////////////////////
		// scenes
    	
    	// get screen resolution
    	Rectangle2D primaryScreenBounds = Screen.getPrimary().getVisualBounds();

		// scene difficulty
		GridPane gridDifficulty = new GridPane();
		gridDifficulty.setAlignment(Pos.CENTER);
		gridDifficulty.setHgap(25);
		gridDifficulty.setVgap(25);
		Scene sceneDifficulty = new Scene(gridDifficulty, 750, 480);

    	// scene mod
        GridPane gridMod = new GridPane();
        gridMod.setAlignment(Pos.CENTER);
        gridMod.setHgap(25);
        gridMod.setVgap(25);
        Scene sceneMod = new Scene(gridMod, 750, 480);

        // scene num of players
        GridPane gridNumOfPlayers = new GridPane();
        gridNumOfPlayers.setAlignment(Pos.CENTER);
        gridNumOfPlayers.setHgap(25);
        gridNumOfPlayers.setVgap(25);
        Scene sceneNumOfPlayers = new Scene(gridNumOfPlayers, 750, 480);
        
		// scene start
		GridPane gridStart = new GridPane();
		gridStart.setAlignment(Pos.CENTER);
		gridStart.setHgap(25);
		gridStart.setVgap(25);
		Scene sceneStart = new Scene(gridStart, 750, 480);
        
		// scene mod i
		GridPane gridGameMod1 = new GridPane();
		gridGameMod1.setAlignment(Pos.CENTER);
		gridGameMod1.setHgap(25);
		gridGameMod1.setVgap(25);
		Scene sceneGameMod1 = new Scene(gridGameMod1, 750, 480);
		
		// scene mod ii
		GridPane gridGameMod2 = new GridPane();
		gridGameMod2.setAlignment(Pos.CENTER);
		gridGameMod2.setHgap(0);
		gridGameMod2.setVgap(0);
		// drawing--
		Canvas canvas = new Canvas(primaryScreenBounds.getHeight(), primaryScreenBounds.getHeight());
        GraphicsContext gc = canvas.getGraphicsContext2D();
        gridGameMod2.add(canvas, 1, 0);
		// --drawing
		Scene sceneGameMod2 = new Scene(gridGameMod2, primaryScreenBounds.getWidth(), primaryScreenBounds.getHeight());
		
		System.out.println("App, scenes created");
		
		////////////////////////////////////////////////////////////////////////
		////////////////////////////////////////////////////////////////////////
		// labels modified by events
		
		Label labelModChoosenDifficulty = new Label();
		gridMod.add(labelModChoosenDifficulty, 0, 0);
		
		Label labelNumChoosenDifficulty = new Label();
        gridNumOfPlayers.add(labelNumChoosenDifficulty, 0, 0);
        
        Label labelStartChoosenDifficulty = new Label();
        gridStart.add(labelStartChoosenDifficulty, 0, 0);
        
        Label labelStart = new Label();
        gridStart.add(labelStart, 0, 1);

        Label labelEnding = new Label();
        gridGameMod2.add(labelEnding, 0, 0);
        
		////////////////////////////////////////////////////////////////////////
		////////////////////////////////////////////////////////////////////////
		// scene difficulty buttons
        
        InterfaceData interfaceData = new InterfaceData();
        
        Label labelDifficulty = new Label("choose difficulty:\n\n"
        		+ "1 - speed race 10x10 3 sec start advantage\n"
        		+ "2 - speed race 15x15 5 sec start advantage\n"
        		+ "3 - speed race 20x20 10 sec start advantage\n"
        		+ "4 - speed race 25x25 20 sec start advantage\n"
        		+ "5 - speed race 30x30 40 sec start advantage\n"
        		+ "6 - speed race 35x35 80 sec start advantage\n"
        		+ "7 - classic 30x30 maze\n"
        		+ "8 - classic 40x40 maze\n"
        		+ "9 - classic 50x50 maze\n");
		gridDifficulty.add(labelDifficulty, 0, 0);

        ChoiceBox<Integer> choiceBox = new ChoiceBox<>();
        choiceBox.getItems().add(1);
        choiceBox.getItems().add(2);
        choiceBox.getItems().add(3);
        choiceBox.getItems().add(4);
        choiceBox.getItems().add(5);
        choiceBox.getItems().add(6);
        choiceBox.getItems().add(7);
        choiceBox.getItems().add(8);
        choiceBox.getItems().add(9);
        
        // defalut value
        choiceBox.setValue(1);
        interfaceData.setInputedDifficulty((Integer)choiceBox.getValue());
        labelModChoosenDifficulty.setText("choosen difficulty is " + (Integer)choiceBox.getValue());
        
        gridDifficulty.add(choiceBox, 0, 1);
        choiceBox.setOnAction(e -> {
        	interfaceData.setInputedDifficulty((Integer)choiceBox.getValue());
        	labelModChoosenDifficulty.setText("choosen difficulty is " + (Integer)choiceBox.getValue());
        });

        Button buttonConfirmDifficulty = new Button("confirm");
        gridDifficulty.add(buttonConfirmDifficulty, 0, 2);
        
        buttonConfirmDifficulty.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent e) {
            	stage.setScene(sceneMod);
            }
        });
        
        System.out.println("App, buttons scene difficulty created");
        
		////////////////////////////////////////////////////////////////////////
		////////////////////////////////////////////////////////////////////////
		// scene mod buttons

        Label labelMod = new Label("pick mod");
        gridMod.add(labelMod, 0, 1, 2, 1);
        
        Button buttonMod1 = new Button("mod 1 - first person view, not yet implemented");
        gridMod.add(buttonMod1, 0, 2);
		
        // not yet implemented
//        buttonMod1.setOnAction(new EventHandler<ActionEvent>() {
//            @Override
//            public void handle(ActionEvent e) {
//            	Static.mod = 1;
//            	stage.setScene(sceneStart);
//            }
//        });
        
        Button buttonMod2 = new Button("mod 2 - top down view, one or two players");
        gridMod.add(buttonMod2, 1, 2);
        
        buttonMod2.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent e) {
            	interfaceData.setMod(2);
            	labelNumChoosenDifficulty.setText("choosen difficulty is " + (Integer)choiceBox.getValue());
            	stage.setScene(sceneNumOfPlayers);
            }
        });
        
        Button buttonModBack = new Button("back");
        gridMod.add(buttonModBack, 0, 3, 2, 1);
        
        buttonModBack.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent e) {
            	stage.setScene(sceneDifficulty);
            }
        });
        
        System.out.println("App, buttons scene 2 created");
        
		////////////////////////////////////////////////////////////////////////
		////////////////////////////////////////////////////////////////////////
		// scene num of players buttons

        Label labelNumOfPlayers = new Label("choose number of players");
        gridNumOfPlayers.add(labelNumOfPlayers, 0, 1, 2, 1);
        
        Button onePlayer = new Button("one");
        gridNumOfPlayers.add(onePlayer, 0, 2);
		
        onePlayer.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent e) {
            	interfaceData.setNumberOfPlayers(1);
            	labelStartChoosenDifficulty.setText("choosen difficulty is " + (Integer)choiceBox.getValue());
            	labelStart.setText("arrows for move, M for marker");
            	stage.setScene(sceneStart);
            }
        });
        
        Button twoPlayers = new Button("two");
        gridNumOfPlayers.add(twoPlayers, 1, 2);
        
        twoPlayers.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent e) {
            	interfaceData.setNumberOfPlayers(2);
            	labelStartChoosenDifficulty.setText("choosen difficulty is " + (Integer)choiceBox.getValue());
            	labelStart.setText("player 1: arrows - move, M - marker,\nplayer 2: ASDW - move B - marker");
            	stage.setScene(sceneStart);
            }
        });
        
        Button buttonNumOfPlayersBack = new Button("back");
        gridNumOfPlayers.add(buttonNumOfPlayersBack, 0, 3, 2, 1);
        
        buttonNumOfPlayersBack.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent e) {
            	stage.setScene(sceneMod);
            }
        });
        
        System.out.println("App, buttons scene 2b created");
        
		////////////////////////////////////////////////////////////////////////
		////////////////////////////////////////////////////////////////////////
		// scene start buttons
        
        Game game = new Game(gc, interfaceData);

        Button start = new Button("start");
        gridStart.add(start, 0, 2);

        start.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent e) {
            	
            	if (interfaceData.getMod() == 1) {
// drawing.drawing(gc); ////////////////////////////////////////////////////////////////////////////////////////////////////
            		stage.setScene(sceneGameMod1);
            	}
            	else if (interfaceData.getMod() == 2) {
            		game.playGame();
                	interfaceData.setGameLiveOrNot(1);
            		game.drawing();
            		stage.setScene(sceneGameMod2);
            		stage.setX((primaryScreenBounds.getWidth() - stage.getWidth()) / 2);
	                stage.setY((primaryScreenBounds.getHeight() - stage.getHeight()) / 2);
            	}
            }
        });
        
        Button buttonStartBack = new Button("back");
        gridStart.add(buttonStartBack, 0, 3);
        
        buttonStartBack.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent e) {
            	if (interfaceData.getMod() == 1) {
            		stage.setScene(sceneMod);
            	} else if (interfaceData.getMod() == 2) {
            		stage.setScene(sceneNumOfPlayers);
            	}
            }
        });
        
        System.out.println("App, buttons scene 3 created");
        
		////////////////////////////////////////////////////////////////////////
		////////////////////////////////////////////////////////////////////////
		// listener commands

        EventHandler<KeyEvent> keyPressHandlerMod1 = new EventHandler<KeyEvent>() {
			@Override
			public void handle(KeyEvent event) {

				if (interfaceData.getMod() == 1) {
					
					System.out.println("App, key events mod 1 evoked a");
					
					if (event.getCode() == KeyCode.UP) {
						Static.buttonFront(game.getMaze(), game.getPlayer());
					}
					else if (event.getCode() == KeyCode.DOWN) {
						Static.buttonBack(game.getMaze(), game.getPlayer());
					}
					else if (event.getCode() == KeyCode.LEFT) {
						Static.leftButton(game.getPlayer());
					}
					else if (event.getCode() == KeyCode.RIGHT) {
						Static.rightButton(game.getPlayer());
					}
					else if (event.getCode() == KeyCode.M) {

						Static.setMarker(game.getJgraphtConvert(), game.getConvertedGraph(), game.getPlayer());
						
						System.out.println("App, marker set");
						
					}
				}
				
				System.out.println("App, key events mod 1 evoked full");
				
// drawing.drawing(gc); ////////////////////////////////////////////////////////////////////////////////////////////////////
				
			}
		};
		sceneGameMod1.setOnKeyPressed(keyPressHandlerMod1);

		EventHandler<KeyEvent> keyPressHandlerMod2 = new EventHandler<KeyEvent>() {
			@Override
			public void handle(KeyEvent event) {
				
				if (event.getCode() == KeyCode.R) {
					
					interfaceData.setGameLiveOrNot(0);

					// clear screen
	            	gc.clearRect(0, 0, canvas.getWidth(), canvas.getHeight());
	            	
	            	labelEnding.setText("");
					
					stage.setScene(sceneDifficulty);

	            	stage.setX((primaryScreenBounds.getWidth() - stage.getWidth()) / 2);
	                stage.setY((primaryScreenBounds.getHeight() - stage.getHeight()) / 2);
					
					System.out.println("app, reset game");
					
				}
				
				if (interfaceData.getMod() == 2 && interfaceData.getGameLiveOrNot() == 1) {

					System.out.println("App, key events mod 2 evoked a");
					
					if(game.getPlayer().getPlayerLife() == 1) {
						
						System.out.println("App, key events mod 2 evoked b");
						
						// up & down switched because of contra drawing on y-axe by javafx
						if (event.getCode() == KeyCode.DOWN) {
							Static.upArrow(game.getMaze(), game.getPlayer());
						}
						else if (event.getCode() == KeyCode.UP) {
							Static.downArrow(game.getMaze(), game.getPlayer());
						}
						else if (event.getCode() == KeyCode.LEFT) {
							Static.leftArrow(game.getMaze(), game.getPlayer());
						}
						else if (event.getCode() == KeyCode.RIGHT) {
							Static.rightArrow(game.getMaze(), game.getPlayer());
						}
						else if (event.getCode() == KeyCode.M) {
							
							Static.setMarker(game.getJgraphtConvert(), game.getConvertedGraph(), game.getPlayer());
							
							System.out.println("App, marker set");
							
						}
						
					}

					System.out.println("App, key events mod 2 evoked full");
					
					if (interfaceData.getNumberOfPlayers() == 2) {
						
						System.out.println("App, key events mod 2 evoked a player 2");

				    	if (game.getPlayer2().getPlayerLife() == 1) {
				    		
				    		// up & down switched because of contra drawing on y-axe by javafx
							if (event.getCode() == KeyCode.S) {
								Static.upArrow(game.getMaze(), game.getPlayer2());
							}
							else if (event.getCode() == KeyCode.W) {
								Static.downArrow(game.getMaze(), game.getPlayer2());
							}
							else if (event.getCode() == KeyCode.A) {
								Static.leftArrow(game.getMaze(), game.getPlayer2());
							}
							else if (event.getCode() == KeyCode.D) {
								Static.rightArrow(game.getMaze(), game.getPlayer2());
							}
							else if (event.getCode() == KeyCode.B) {
								
								Static.setMarker(game.getJgraphtConvert(), game.getConvertedGraph(), game.getPlayer2());
								
								System.out.println("App, marker set");
								
							}
				    		
				    	}

						System.out.println("App, key events mod 2 evoked full player 2");
						
					}
					
					game.drawing();
					
					// check for ending
					ending(interfaceData, game, labelEnding);

				}
			}
		};
		sceneGameMod2.setOnKeyPressed(keyPressHandlerMod2);

		////////////////////////////////////////////////////////////////////////
		////////////////////////////////////////////////////////////////////////
		// redraw end ending
		
		// animation timer class update drawing.. slowed down to 0.05 seconds
		new AnimationTimer() {
			private long lastUpdate = 0 ;
			@Override
			public void handle(long now) {
				if(interfaceData.getGameLiveOrNot() == 1) {
					// 50 miliseconds in nanoseconds
					if (now - lastUpdate >= 50_000_000) {
						
						// functions to execute
						game.drawing();
						ending(interfaceData, game, labelEnding);
						
	                    lastUpdate = now ;
	                }
				}
			}
        }.start();
        
        // close application on last closed stage
        stage.setOnCloseRequest(new EventHandler<WindowEvent>() {
            @Override
            public void handle(WindowEvent t) {
                Platform.exit();
                System.exit(0);
            }
        });

        stage.setTitle("amaze");
        stage.setScene(sceneDifficulty); // default scene
        stage.show();
        
    }

    public void ending(InterfaceData interfaceData, Game game, Label labelEnding) {
    	// check for ending
		if (interfaceData.getNumberOfPlayers() == 1) {
			if (game.getPlayer().getPlayerLife() == 0) {
				interfaceData.setGameLiveOrNot(0);
				labelEnding.setText("MORE LUCK NEXT TIME\n\nPUSH R TO RESET");
			} else if (game.getPlayer().getPlayerPositionX() == game.getExitNode().getPosX() && game.getPlayer().getPlayerPositionY() == game.getExitNode().getPosY()) {
				interfaceData.setGameLiveOrNot(0);
				labelEnding.setText("CONGRATULATIONS\n YOU WIN\n\nPUSH R TO RESET");
			}
		} else if (interfaceData.getNumberOfPlayers() == 2) {
			if (game.getPlayer().getPlayerLife() == 0 && game.getPlayer2().getPlayerLife() == 0) {
				interfaceData.setGameLiveOrNot(0);
				labelEnding.setText("YOU BOTH SUCK\n\nPUSH R TO RESET");
			} else if (game.getPlayer().getPlayerPositionX() == game.getExitNode().getPosX() && game.getPlayer().getPlayerPositionY() == game.getExitNode().getPosY()) {
				interfaceData.setGameLiveOrNot(0);
				labelEnding.setText("CONGRATULATIONS PLAYER 1\n YOU ROCK\n\nPUSH R TO RESET");
			} else if (game.getPlayer2().getPlayerPositionX() == game.getExitNode().getPosX() && game.getPlayer2().getPlayerPositionY() == game.getExitNode().getPosY()) {
				interfaceData.setGameLiveOrNot(0);
				labelEnding.setText("CONGRATULATIONS PLAYER 2\n YOU ROCK\n\nPUSH R TO RESET");
			}
		}
    }
    
    public static void main(String[] args) {
        launch();
    }

}