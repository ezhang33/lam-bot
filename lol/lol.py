# game.py
import pygame, sys
from map import GameMap
from lol.champion import Champion
from minion import Minion

pygame.init()
SCREEN = pygame.display.set_mode((800, 600))
CLOCK  = pygame.time.Clock()

def main():
    game_map = GameMap("assets/map.tmx")       # e.g. Tiled map
    
    player = Champion((100,100))
    minions = [Minion((400,300)) for _ in range(5)]

    while True:
        dt = CLOCK.tick(60) / 1000.0

        # input
        for e in pygame.event.get():
            if e.type == pygame.QUIT:
                pygame.quit(); sys.exit()
            player.handle_event(e)

        # update
        player.update(dt)

        # collision: projectiles ↔ minions
        for p in player.projectiles[:]:
            for m in minions:
                if (p.pos - m.pos).length() < p.radius + 10:
                    m.take_damage(p.damage)
                    p.alive = False
                    break

        # clean up dead minions
        minions = [m for m in minions if m.alive]

        # champion auto-attack on close minions
        for m in minions:
            if (player.pos - m.pos).length() < player.attack_range:
                player.attack(m)

        # render
        SCREEN.fill((50,50,60))
        game_map.draw(SCREEN)
        for m in minions: m.draw(SCREEN)
        player.draw(SCREEN)
        pygame.display.flip()

if __name__ == "__main__":
    main()
