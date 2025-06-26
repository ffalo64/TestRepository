import tkinter as tk
import random

TILE_SIZE = 10
MAP_WIDTH = 50
MAP_HEIGHT = 50
FLOORS = 5

WALL = '#'
FLOOR = '.'
STAIRS = '>'
PLAYER = '@'

class Room:
    def __init__(self, x, y, w, h):
        self.x1 = x
        self.y1 = y
        self.x2 = x + w
        self.y2 = y + h

    def center(self):
        cx = (self.x1 + self.x2) // 2
        cy = (self.y1 + self.y2) // 2
        return cx, cy

    def intersects(self, other):
        return (self.x1 <= other.x2 and self.x2 >= other.x1 and
                self.y1 <= other.y2 and self.y2 >= other.y1)

class Dungeon:
    def __init__(self):
        self.floor = 1
        self.maps = []
        self.player_pos = (0, 0)
        for i in range(FLOORS):
            game_map, stairs_pos, start_pos = self.generate_floor(i == FLOORS - 1)
            self.maps.append({'map': game_map,
                              'stairs': stairs_pos,
                              'start': start_pos})
            if i == 0:
                self.player_pos = start_pos

    def generate_floor(self, is_last=False):
        game_map = [[WALL for _ in range(MAP_WIDTH)] for _ in range(MAP_HEIGHT)]
        rooms = []
        max_rooms = 10
        min_size = 4
        max_size = 10
        for _ in range(max_rooms):
            w = random.randint(min_size, max_size)
            h = random.randint(min_size, max_size)
            x = random.randint(1, MAP_WIDTH - w - 1)
            y = random.randint(1, MAP_HEIGHT - h - 1)
            new_room = Room(x, y, w, h)
            if any(new_room.intersects(r) for r in rooms):
                continue
            for i in range(new_room.x1, new_room.x2):
                for j in range(new_room.y1, new_room.y2):
                    game_map[j][i] = FLOOR
            if rooms:
                prev_x, prev_y = rooms[-1].center()
                new_x, new_y = new_room.center()
                if random.random() < 0.5:
                    for x_ in range(min(prev_x, new_x), max(prev_x, new_x)+1):
                        game_map[prev_y][x_] = FLOOR
                    for y_ in range(min(prev_y, new_y), max(prev_y, new_y)+1):
                        game_map[y_][new_x] = FLOOR
                else:
                    for y_ in range(min(prev_y, new_y), max(prev_y, new_y)+1):
                        game_map[y_][prev_x] = FLOOR
                    for x_ in range(min(prev_x, new_x), max(prev_x, new_x)+1):
                        game_map[new_y][x_] = FLOOR
            rooms.append(new_room)
        start_x, start_y = rooms[0].center()
        stair_x, stair_y = rooms[-1].center()
        game_map[stair_y][stair_x] = STAIRS if not is_last else 'X'
        return game_map, (stair_x, stair_y), (start_x, start_y)

    def current_map(self):
        return self.maps[self.floor - 1]['map']

    def stairs_pos(self):
        return self.maps[self.floor - 1]['stairs']

    def move_player(self, dx, dy):
        x, y = self.player_pos
        nx, ny = x + dx, y + dy
        if nx < 0 or nx >= MAP_WIDTH or ny < 0 or ny >= MAP_HEIGHT:
            return
        tile = self.current_map()[ny][nx]
        if tile == WALL:
            return
        self.player_pos = (nx, ny)
        if (nx, ny) == self.stairs_pos():
            if self.floor == FLOORS:
                return 'clear'
            else:
                self.floor += 1
                self.player_pos = self.maps[self.floor - 1]['start']

class GameGUI:
    def __init__(self, root, dungeon):
        self.root = root
        self.dungeon = dungeon
        canvas_size = TILE_SIZE * MAP_WIDTH
        self.canvas = tk.Canvas(root, width=canvas_size, height=canvas_size)
        self.canvas.pack()
        root.bind('<Key>', self.on_key)
        self.draw_map()

    def draw_map(self):
        self.canvas.delete('all')
        game_map = self.dungeon.current_map()
        px, py = self.dungeon.player_pos
        for y, row in enumerate(game_map):
            for x, tile in enumerate(row):
                color = 'black'
                if tile == WALL:
                    color = 'gray'
                elif tile == FLOOR:
                    color = 'white'
                elif tile == STAIRS:
                    color = 'blue'
                elif tile == 'X':
                    color = 'green'
                self.canvas.create_rectangle(x*TILE_SIZE, y*TILE_SIZE,
                                             (x+1)*TILE_SIZE, (y+1)*TILE_SIZE,
                                             fill=color, outline='')
        self.canvas.create_rectangle(px*TILE_SIZE, py*TILE_SIZE,
                                     (px+1)*TILE_SIZE, (py+1)*TILE_SIZE,
                                     fill='red', outline='')

    def on_key(self, event):
        key = event.keysym
        move = {
            'Up': (0, -1),
            'Down': (0, 1),
            'Left': (-1, 0),
            'Right': (1, 0)
        }
        if key in move:
            result = self.dungeon.move_player(*move[key])
            if result == 'clear':
                self.canvas.delete('all')
                self.canvas.create_text(250, 250, text='Dungeon Cleared!', font=('Arial', 24))
                return
            self.draw_map()

if __name__ == '__main__':
    random.seed()
    d = Dungeon()
    root = tk.Tk()
    root.title('Roguelike')
    game = GameGUI(root, d)
    root.mainloop()
