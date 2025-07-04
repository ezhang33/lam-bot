# champion.py
import pygame
from pygame.math import Vector2

class Ability:
    def __init__(self, name, key, cooldown, cast_fn):
        self.name = name
        self.key = key
        self.cooldown = cooldown
        self._timer = 0.0
        self.cast_fn = cast_fn

    def ready(self):
        return self._timer <= 0.0

    def trigger(self, *args, **kwargs):
        if self.ready():
            self.cast_fn(*args, **kwargs)
            self._timer = self.cooldown

    def update(self, dt):
        self._timer = max(0.0, self._timer - dt)

class Projectile:
    def __init__(self, pos, direction, speed=400, damage=50, radius=5):
        self.pos = Vector2(pos)
        self.dir = direction.normalize()
        self.speed = speed
        self.damage = damage
        self.radius = radius
        self.alive = True

    def update(self, dt):
        self.pos += self.dir * self.speed * dt
        # simple bounds‐check (screen 800×600)
        if not (0 <= self.pos.x <= 800 and 0 <= self.pos.y <= 600):
            self.alive = False

    def draw(self, surf):
        pygame.draw.circle(surf, (255,200,0), self.pos, self.radius)

class Champion:
    def __init__(self, pos):
        self.pos = Vector2(pos)
        self.speed = 200
        # auto-attack
        self.attack_range = 75
        self.attack_dmg   = 20
        self.attack_cd    = 1.0
        self._atk_timer   = 0.0
        # status effects
        self.shield_timer = 0.0
        self.buff_timer   = 0.0
        # projectile list
        self.projectiles  = []

        # map QWER to ability objects
        self.abilities = {
            pygame.K_q: Ability("Q Fireball", pygame.K_q, cooldown=2.0, cast_fn=self.cast_q),
            pygame.K_w: Ability("W Shield",    pygame.K_w, cooldown=8.0, cast_fn=self.cast_w),
            pygame.K_e: Ability("E Dash",      pygame.K_e, cooldown=5.0, cast_fn=self.cast_e),
            pygame.K_r: Ability("R Empower",   pygame.K_r, cooldown=20.0, cast_fn=self.cast_r),
        }

    def handle_event(self, evt):
        if evt.type == pygame.MOUSEBUTTONDOWN and evt.button == 1:
            self.move_target = Vector2(evt.pos)
        elif evt.type == pygame.KEYDOWN and evt.key in self.abilities:
            self.abilities[evt.key].trigger()

    def update(self, dt):
        # movement toward click
        if hasattr(self, "move_target"):
            dir = (self.move_target - self.pos)
            if dir.length() > 5:
                self.pos += dir.normalize() * self.speed * dt

        # auto-attack cooldown
        self._atk_timer = max(0.0, self._atk_timer - dt)
        # ability cooldowns
        for ab in self.abilities.values():
            ab.update(dt)
        # status timers
        if self.shield_timer > 0:
            self.shield_timer = max(0.0, self.shield_timer - dt)
        if self.buff_timer > 0:
            self.buff_timer = max(0.0, self.buff_timer - dt)
            if self.buff_timer == 0:
                self.attack_dmg /= 1.5  # remove buff

        # update projectiles
        for p in self.projectiles:
            p.update(dt)
        # remove dead projectiles
        self.projectiles = [p for p in self.projectiles if p.alive]

    def attack(self, target):
        if self._atk_timer == 0.0:
            dmg = self.attack_dmg
            # shield absorbs half damage if active
            if self.shield_timer > 0:
                dmg *= 0.5
            target.take_damage(dmg)
            self._atk_timer = self.attack_cd

    def draw(self, surf):
        # blue circle if shielded, green if normal
        color = (100,100,255) if self.shield_timer > 0 else (100,200,100)
        pygame.draw.circle(surf, color, self.pos, 16)
        # draw projectiles
        for p in self.projectiles:
            p.draw(surf)

    # --- Ability implementations ---

    def cast_q(self):
        """Fire a projectile toward the mouse."""
        mouse = Vector2(pygame.mouse.get_pos())
        dir   = mouse - self.pos
        if dir.length() > 0:
            self.projectiles.append(
                Projectile(self.pos, dir, speed=400, damage=50)
            )

    def cast_w(self):
        """Gain a 3-second shield that halves incoming damage."""
        self.shield_timer = 3.0

    def cast_e(self):
        """Dash 120 pixels toward the mouse."""
        mouse = Vector2(pygame.mouse.get_pos())
        dir   = mouse - self.pos
        if dir.length() > 0:
            self.pos += dir.normalize() * 120

    def cast_r(self):
        """Empower: 50% increased auto-attack damage for 5 seconds."""
        self.buff_timer  = 5.0
        self.attack_dmg *= 1.5
