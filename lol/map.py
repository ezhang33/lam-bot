# map.py
import pygame

class GameMap:
    def __init__(self, map_file):
        # load Tiled TMX or simple background
        self.bg = pygame.Surface((1600,1200))
        self.bg.fill((30,30,40))

    def draw(self, surf):
        surf.blit(self.bg, (-200, -150))  # simple offset
