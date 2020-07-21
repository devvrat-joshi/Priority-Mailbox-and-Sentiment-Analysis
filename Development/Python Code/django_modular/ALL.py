import pygame 
  
# activate the pygame library . 
# initiate pygame and give permission 
# to use pygame's functionality. 
pygame.init() 
  
# define the RGB value 
# for white colour 
white = (255, 255, 255) 

# assigning values to X and Y variable 
X = 600
Y = 400
  
# create the display surface object 
# of specific dimension..e(X, Y). 
display_surface = pygame.display.set_mode((X, Y )) 
green = (0, 255, 0) 
blue = (255, 255, 255)
# set the pygame window name 
pygame.display.set_caption('Modular Addin Installer') 
font = pygame.font.Font('freesansbold.ttf', 26) 
  
# create a text suface object, 
# on which text is drawn on it. 
text = font.render('Modular Addin Installer', True, green, None) 
text2 = font.render('Install', True, (0,0,255),None) 
# create a rectangular object for the 
# text surface object 
textRect = text.get_rect()  
textRect2 = text2.get_rect()
textRect2.center = (540,340)
# set the center of the rectangular object. 
textRect.center = (X // 2, Y // 6) 
# create a surface object, image is drawn on it. 
image = pygame.image.load("final.jpg") 
  
# infinite loop 
while True : 
  
    # completely fill the surface object 
    # with white colour 
    display_surface.fill(white) 
    
    # copying the image surface object 
    # to the display surface object at 
    # (0, 0) coordinate. 

    display_surface.blit(image, (0, 0)) 
    display_surface.blit(text, textRect)     

    
    # iterate over the list of Event objects 
    # that was returned by pygame.event.get() method. 
    for event in pygame.event.get() : 
  
        # if event object type is QUIT 
        # then quitting the pygame 
        # and program both. 
        if event.type == pygame.QUIT : 
  
            # deactivates the pygame library 
            pygame.quit() 
  
            # quit the program. 
            quit() 
    mouse = pygame.mouse.get_pos()
    x,y = mouse
    
    if(580>x and x>500 and y>320 and y<360):
        print("this")
        pygame.draw.rect(display_surface,(0,255,0),(500,320,80,40))
    else:
        print(x,y)
        pygame.draw.rect(display_surface,(240,240,240),(500,320,80,40)) 
    display_surface.blit(text2, textRect2) 
        # Draws the surface object to the screen.   
    pygame.display.update()  