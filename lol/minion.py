# minion.py
import pygame
from pygame.math import Vector2

class Minion:
    def __init__(self, pos):
        self.pos   = Vector2(pos)
        self.hp    = 50
        self.alive = True

    def take_damage(self, amount):
        self.hp -= amount
        if self.hp <= 0:
            self.alive = False

    def update(self, dt):
        pass  # you could add walking or targeting logic here

    def draw(self, surf):
        pygame.draw.rect(surf, (200,100,100), (*self.pos-Vector2(10,10), 20,20))
