import win32com.client
import pythoncom
import time
import math
import random
import numpy as np


class CATIAV5Controller:
    def __init__(self):
        '''
        #随机生成图形
        pythoncom.CoInitialize()
        try:
            self.catia = win32com.client.Dispatch("CATIA.Application")
            self.catia.Visible = True
            print("=== CATIA V5 自动化绘图 ===")
            print("CATIA 连接成功")

            # 随机生成圆形和矩形，确保初始不重叠
            self.shapes = self.generate_non_overlapping_shapes(shape_count=random.randint(3, 8))

            self.best_solution = None
            self.sketch = None
            self.part = None
            self.part_doc = None
            self.main_body = None
        except Exception as e:
            print(f"初始化失败: {e}")
            self._cleanup()
            raise
        '''

        pythoncom.CoInitialize()
        try:
            self.catia = win32com.client.Dispatch("CATIA.Application")
            self.catia.Visible = True
            print("=== CATIA V5 自动化绘图 ===")
            print("CATIA 连接成功")

            # 修改为：1个圆形 + 4个正方形
            self.shapes = [
                {
                    'type': 'circle',
                    'x': 100,
                    'y': 100,
                    'radius': 8,
                    'fixed_size': True
                },
                {
                    'type': 'rectangle',
                    'x': 120,
                    'y': 100,
                    'width': 10,
                    'height': 10,
                    'angle': 0,
                    'fixed_size': True
                },
                {
                    'type': 'rectangle',
                    'x': 80,
                    'y': 100,
                    'width': 10,
                    'height': 10,
                    'angle': math.pi / 6,
                    'fixed_size': False
                },
                {
                    'type': 'rectangle',
                    'x': 100,
                    'y': 120,
                    'width': 10,
                    'height': 10,
                    'angle': math.pi / 4,
                    'fixed_size': False
                },
                {
                    'type': 'rectangle',
                    'x': 100,
                    'y': 80,
                    'width': 10,
                    'height': 10,
                    'angle': math.pi / 3,
                    'fixed_size': False
                }
            ]

            self.best_solution = None
            self.sketch = None
            self.part = None
            self.part_doc = None
            self.main_body = None
        except Exception as e:
            print(f"初始化失败: {e}")
            self._cleanup()
            raise

    def generate_non_overlapping_shapes(self, shape_count=5, max_attempts=100):
        """随机生成不重叠的圆形和矩形"""
        shapes = []
        attempts = 0

        while len(shapes) < shape_count and attempts < max_attempts:
            # 随机决定生成圆形还是矩形
            if random.random() < 0.5:  # 50%概率生成圆形
                new_shape = {
                    'type': 'circle',
                    'x': random.uniform(80, 120),
                    'y': random.uniform(80, 120),
                    'radius': random.uniform(5, 15),
                    'fixed_size': False  # 默认不是固定大小的
                }
                # 第一个圆形设为插头(固定大小)
                if not any(s['type'] == 'circle' for s in shapes):
                    new_shape['radius'] = 10  # 固定半径
                    new_shape['fixed_size'] = True
            else:  # 生成矩形
                new_shape = {
                    'type': 'rectangle',
                    'width': random.uniform(8, 15),
                    'height': random.uniform(6, 12),
                    'x': random.uniform(80, 120),
                    'y': random.uniform(80, 120),
                    'angle': random.uniform(-math.pi / 2, math.pi / 2),
                    'fixed_size': False
                }

            # 检查新形状是否与现有形状重叠
            overlaps = False
            for existing_shape in shapes:
                if self.check_overlap(new_shape, existing_shape):
                    overlaps = True
                    break

            if not overlaps:
                shapes.append(new_shape)
            attempts += 1

        if len(shapes) < shape_count:
            print(f"警告: 只生成了 {len(shapes)} 个不重叠的形状，未能达到 {shape_count} 个")

        return shapes

    def _cleanup(self):
        """清理资源"""
        try:
            if hasattr(self, 'catia') and self.catia:
                self.catia.Quit()
        except Exception as e:
            print(f"清理资源时出错: {e}")
        finally:
            pythoncom.CoUninitialize()

    def hide_reference_planes(self):
        """隐藏XY/YZ/ZX基准平面"""
        try:
            origin_elements = self.part.OriginElements
            origin_elements.PlaneXY.Visible = False
            origin_elements.PlaneYZ.Visible = False
            origin_elements.PlaneZX.Visible = False
            self.part.Update()
            print("已隐藏XY/YZ/ZX基准平面")
            return True
        except Exception as e:
            print(f"隐藏基准平面失败: {e}")
            return False

    def create_part(self):
        """创建零件并确保 PartBody 存在"""
        try:
            self.part_doc = self.catia.Documents.Add("Part")
            self.part = self.part_doc.Part
            self.bodies = self.part.Bodies

            try:
                self.main_body = self.bodies.Item("PartBody")
            except:
                try:
                    self.main_body = self.bodies.Item("零件体")
                except:
                    print("未找到 PartBody，创建新的 Body...")
                    self.main_body = self.bodies.Add()
                    self.main_body.Name = "PartBody"
                    print("PartBody 创建成功")

            self.hide_reference_planes()
            return True
        except Exception as e:
            print(f"零件创建失败: {e}")
            self._cleanup()
            raise

    def create_sketch(self, plane="xy"):
        """创建草图"""
        try:
            origin_elements = self.part.OriginElements
            plane_obj = {
                "xy": origin_elements.PlaneXY,
                "yz": origin_elements.PlaneYZ,
                "zx": origin_elements.PlaneZX
            }.get(plane.lower(), origin_elements.PlaneXY)

            sketches = self.main_body.Sketches
            sketch = sketches.Add(plane_obj)
            self.part.InWorkObject = sketch
            self.part.Update()
            print("草图创建成功")
            return sketch
        except Exception as e:
            print(f"草图创建失败: {e}")
            self._cleanup()
            raise

    def draw_circle(self, sketch, x=0, y=0, radius=50, name=None):
        """绘制圆形"""
        try:
            factory = sketch.OpenEdition()
            circle = factory.CreateClosedCircle(x, y, radius)
            if name:
                circle.Name = name
            sketch.CloseEdition()
            self.part.Update()
            return True
        except Exception as e:
            print(f"绘制圆形失败: {e}")
            return False

    def draw_rectangle(self, sketch, x, y, width, height, angle=0):
        """绘制旋转矩形"""
        try:
            factory = sketch.OpenEdition()
            half_w = width / 2
            half_h = height / 2

            cos_a = math.cos(angle)
            sin_a = math.sin(angle)

            points = [
                (-half_w, -half_h),
                (half_w, -half_h),
                (half_w, half_h),
                (-half_w, half_h),
                (-half_w, -half_h)
            ]

            rotated_points = [
                (x + px * cos_a - py * sin_a,
                 y + px * sin_a + py * cos_a)
                for px, py in points
            ]

            for i in range(4):
                factory.CreateLine(
                    rotated_points[i][0], rotated_points[i][1],
                    rotated_points[i + 1][0], rotated_points[i + 1][1]
                )

            sketch.CloseEdition()
            self.part.Update()
            print(
                f"矩形创建成功 (中心: ({x:.2f}, {y:.2f}), 宽: {width:.2f}, 高: {height:.2f}, 角度: {math.degrees(angle):.1f}°)")
            return True
        except Exception as e:
            print(f"绘制矩形失败: {e}")
            return False

    def get_shape_corners(self, shape):
        """获取形状的所有角点坐标"""
        if shape['type'] == 'rectangle':
            x, y = shape['x'], shape['y']
            half_w = shape['width'] / 2
            half_h = shape['height'] / 2
            angle = shape['angle']

            cos_a = math.cos(angle)
            sin_a = math.sin(angle)

            corners = []
            for dx, dy in [(-half_w, -half_h), (half_w, -half_h),
                           (half_w, half_h), (-half_w, half_h)]:
                rot_x = x + dx * cos_a - dy * sin_a
                rot_y = y + dx * sin_a + dy * cos_a
                corners.append((rot_x, rot_y))
            return corners
        elif shape['type'] == 'circle':
            # 对于圆形，返回圆周上的8个点作为近似
            x, y = shape['x'], shape['y']
            radius = shape['radius']
            corners = []
            for i in range(8):
                angle = 2 * math.pi * i / 8
                corners.append((x + radius * math.cos(angle), y + radius * math.sin(angle)))
            return corners

    def check_overlap(self, shape1, shape2):
        """检查两个形状是否重叠(分离轴定理)"""

        def project(poly, axis):
            dots = [p[0] * axis[0] + p[1] * axis[1] for p in poly]
            return min(dots), max(dots)

        corners1 = self.get_shape_corners(shape1)
        corners2 = self.get_shape_corners(shape2)

        # 检查形状1的边
        edges = []
        for i in range(len(corners1)):
            x1, y1 = corners1[i]
            x2, y2 = corners1[(i + 1) % len(corners1)]
            edge = (x2 - x1, y2 - y1)
            length = math.sqrt(edge[0] ** 2 + edge[1] ** 2)
            if length > 0:
                normal = (-edge[1] / length, edge[0] / length)
                edges.append(normal)

        # 检查形状2的边
        for i in range(len(corners2)):
            x1, y1 = corners2[i]
            x2, y2 = corners2[(i + 1) % len(corners2)]
            edge = (x2 - x1, y2 - y1)
            length = math.sqrt(edge[0] ** 2 + edge[1] ** 2)
            if length > 0:
                normal = (-edge[1] / length, edge[0] / length)
                edges.append(normal)

        # 在所有分离轴上投影
        for axis in edges:
            min1, max1 = project(corners1, axis)
            min2, max2 = project(corners2, axis)

            if max1 < min2 or max2 < min1:
                return False  # 有分离轴，不重叠

        return True  # 所有轴上都重叠

    def get_bounding_circle(self, shapes):
        """同时考虑矩形角点和圆形外边界来拟合最小包围圆（近似）"""
        all_points = []

        for shape in shapes:
            if shape['type'] == 'rectangle':
                all_points.extend(self.get_shape_corners(shape))
            elif shape['type'] == 'circle':
                # 将圆的“外围边界点”加入考虑
                cx, cy, r = shape['x'], shape['y'], shape['radius']
                # 这里取4个方向（你也可以用8或更多方向更精细）
                all_points.append((cx + r, cy))
                all_points.append((cx - r, cy))
                all_points.append((cx, cy + r))
                all_points.append((cx, cy - r))

        if not all_points:
            return {'x': 0, 'y': 0, 'radius': 0}

        # 求点集中心
        avg_x = sum(p[0] for p in all_points) / len(all_points)
        avg_y = sum(p[1] for p in all_points) / len(all_points)

        # 半径 = 所有点到中心的最大距离
        max_r = max(math.hypot(p[0] - avg_x, p[1] - avg_y) for p in all_points)

        return {'x': avg_x, 'y': avg_y, 'radius': max_r}

    def evaluate_solution(self, solution):
        """评估解决方案的质量 - 改进版"""
        circle_x, circle_y, circle_r = solution['circle']
        shapes = solution['shapes']

        # 检查形状间是否重叠
        overlap_penalty = 0
        for i in range(len(shapes)):
            for j in range(i + 1, len(shapes)):
                if self.check_overlap(shapes[i], shapes[j]):
                    overlap_penalty += 1000  # 大惩罚项

        # 检查所有形状是否在圆内，并计算紧凑度
        max_distance = 0
        total_distance = 0
        shape_areas = []

        for shape in shapes:
            corners = self.get_shape_corners(shape)
            for corner_x, corner_y in corners:
                distance = math.sqrt((corner_x - circle_x) ** 2 + (corner_y - circle_y) ** 2)
                if distance > max_distance:
                    max_distance = distance
                total_distance += distance

            # 计算形状面积用于权重
            if shape['type'] == 'rectangle':
                shape_areas.append(shape['width'] * shape['height'])
            else:
                shape_areas.append(math.pi * shape['radius'] ** 2)

        # 计算紧凑度指标 (考虑形状面积权重)
        avg_distance = total_distance / (len(shapes) * 4)  # 每个形状4个角点
        compactness = sum(area * avg_distance for area in shape_areas) / sum(shape_areas)

        # 适应度：圆半径 + 紧凑度 + 重叠惩罚
        if max_distance <= circle_r:
            return circle_r + 0.5 * compactness + overlap_penalty  # 权重可调
        else:
            return float('inf')  # 无效解

    def simulated_annealing(self, iterations=1000, temp=1000, cooling_rate=0.99):
        """模拟退火优化形状位置和包围圆 - 改进版"""
        # 初始化当前解
        current_solution = {
            'circle': [100, 100, 50],  # 初始圆参数[x, y, radius]
            'shapes': [s.copy() for s in self.shapes]  # 复制初始形状
        }

        # 初始化最佳解
        best_solution = {
            'circle': current_solution['circle'].copy(),
            'shapes': [s.copy() for s in current_solution['shapes']]
        }
        best_fitness = self.evaluate_solution(best_solution)

        for i in range(iterations):
            # 动态调整移动步长(随着温度降低而减小)
            move_step = max(1, 5 * temp / 1000)
            rotate_step = max(0.05, 0.2 * temp / 1000)
            radius_step = max(0.1, 1 * temp / 1000)

            # 生成新解 - 深拷贝当前解
            new_solution = {
                'circle': current_solution['circle'].copy(),
                'shapes': [s.copy() for s in current_solution['shapes']]
            }

            # 随机调整每个形状
            for shape in new_solution['shapes']:
                # 随机平移(步长随温度降低)
                shape['x'] += random.uniform(-move_step, move_step)
                shape['y'] += random.uniform(-move_step, move_step)

                # 如果是矩形，随机旋转
                if shape['type'] == 'rectangle':
                    shape['angle'] += random.uniform(-rotate_step, rotate_step)

                '''
                # 如果是圆形且不是固定大小的，随机调整半径
                if shape['type'] == 'circle' and not shape['fixed_size']:
                    shape['radius'] = max(3, shape['radius'] + random.uniform(-radius_step, radius_step))
                '''

            # 调整圆参数(动态步长)
            '''
            new_solution['circle'][0] += random.uniform(-move_step, move_step)  # x
            new_solution['circle'][1] += random.uniform(-move_step, move_step)  # y
            new_solution['circle'][2] = max(10, new_solution['circle'][2] + random.uniform(-radius_step, radius_step))  # 半径
            '''

            # 自动拟合当前图形的最小外包圆
            bounding = self.get_bounding_circle(new_solution['shapes'])
            new_solution['circle'] = [bounding['x'], bounding['y'], bounding['radius']]

            # 计算适应度
            current_fitness = self.evaluate_solution(current_solution)
            new_fitness = self.evaluate_solution(new_solution)

            # 决定是否接受新解
            if new_fitness < current_fitness:
                accept = True
            else:
                delta = new_fitness - current_fitness
                probability = math.exp(-delta / temp)
                accept = random.random() < probability

            if accept:
                current_solution = new_solution
                current_fitness = new_fitness

                if new_fitness < best_fitness:
                    best_solution = {
                        'circle': new_solution['circle'].copy(),
                        'shapes': [s.copy() for s in new_solution['shapes']]
                    }
                    best_fitness = new_fitness

            # 降低温度
            temp *= cooling_rate

            # 打印进度
            if i % 100 == 0:
                print(
                    f"Iteration {i}: Temp={temp:.2f}, Best Radius={best_solution['circle'][2]:.2f}, Fitness={best_fitness:.2f}")

        return best_solution

    def draw_solution(self, solution):
        """绘制解决方案"""
        try:
            # 先重置草图，确保清除所有旧图形
            if not self.reset_sketch():
                raise Exception("无法重置草图")

            # 绘制所有形状
            for i, shape in enumerate(solution['shapes']):
                if shape['type'] == 'rectangle':
                    if not self.draw_rectangle(self.sketch, shape['x'], shape['y'],
                                               shape['width'], shape['height'], shape['angle']):
                        raise Exception(f"无法绘制第 {i + 1} 个矩形")
                elif shape['type'] == 'circle':
                    if not self.draw_circle(self.sketch, shape['x'], shape['y'], shape['radius']):
                        raise Exception(f"无法绘制第 {i + 1} 个圆形")

            # 绘制包围圆
            cx, cy, cr = solution['circle']
            if not self.draw_circle(self.sketch, cx, cy, cr):
                raise Exception("无法绘制包围圆")

            return True
        except Exception as e:
            print(f"绘制解决方案失败: {e}")
            return False

    def reset_sketch(self, plane="xy"):
        """删除当前草图并新建一个新的草图"""
        try:
            # 获取当前活动文档
            active_doc = self.catia.ActiveDocument

            # 删除所有现有草图
            bodies = self.part.Bodies
            for i in range(1, bodies.Count + 1):
                body = bodies.Item(i)
                sketches = body.Sketches
                for j in range(1, sketches.Count + 1):
                    try:
                        sketch = sketches.Item(j)
                        sketch.Delete()
                    except:
                        continue

            # 创建新的草图
            origin_elements = self.part.OriginElements
            plane_obj = {
                "xy": origin_elements.PlaneXY,
                "yz": origin_elements.PlaneYZ,
                "zx": origin_elements.PlaneZX
            }.get(plane.lower(), origin_elements.PlaneXY)

            sketches = self.main_body.Sketches
            self.sketch = sketches.Add(plane_obj)
            self.part.InWorkObject = self.sketch
            self.part.Update()

            print("草图已重置")
            return True
        except Exception as e:
            print(f"重置草图失败: {e}")
            return False

    def toggle_visibility(self, name, visible=True):
        """控制图形可见性"""
        try:
            for body in self.part.Bodies:
                for sketch in body.Sketches:
                    for element in sketch.Elements:
                        if element.Name == name:
                            element.Visible = visible
            self.part.Update()
            return True
        except Exception as e:
            print(f"控制可见性失败: {e}")
            return False

    def shift_solution_to_target_center(self, solution, target_x, target_y):
        current_cx, current_cy, _ = solution['circle']
        dx = target_x - current_cx
        dy = target_y - current_cy

        # 移动外包圆
        solution['circle'][0] += dx
        solution['circle'][1] += dy

        # 移动所有图形
        for shape in solution['shapes']:
            shape['x'] += dx
            shape['y'] += dy



    def run(self):
        """运行主程序"""
        try:
            if not self.create_part():
                return False

            # 1. 绘制初始布局
            print("\n绘制初始布局...")
            self.sketch = self.create_sketch()
            initial_shapes_sketch = self.sketch.Name  # 保存初始草图名称

            for shape in self.shapes:
                if shape['type'] == 'rectangle':
                    self.draw_rectangle(self.sketch, shape['x'], shape['y'],
                                        shape['width'], shape['height'], shape['angle'])
                elif shape['type'] == 'circle':
                    self.draw_circle(self.sketch, shape['x'], shape['y'], shape['radius'])

            # 2. 优化布局
            input("\n按回车键开始优化布局...")
            print("\n正在优化布局(可能需要几分钟)...")
            self.best_solution = self.simulated_annealing(iterations=2000)

            self.shift_solution_to_target_center(self.best_solution, -100, -100)

            # 3. 绘制优化结果到新草图
            print("\n绘制优化后的布局...")
            if not self.draw_solution(self.best_solution):
                return False

            # 可选: 隐藏初始草图
            self.toggle_visibility(initial_shapes_sketch, visible=False)

            # 显示最终结果
            cx, cy, cr = self.best_solution['circle']
            print(f"\n优化完成:")
            print(f"最小包围圆: 中心({cx:.2f}, {cy:.2f}), 半径: {cr:.2f}")
            print(f"最小包围圆直径: {cr * 2:.2f}")

            input("\n操作完成！按 Enter 退出...")
            return True
        except Exception as e:
            print(f"运行失败: {e}")
            return False


if __name__ == "__main__":
    app = None
    try:
        app = CATIAV5Controller()
        time.sleep(2)
        if not app.run():
            print("执行失败，请检查错误信息")
    except Exception as e:
        print(f"程序运行出错: {e}")
    finally:
        if app:
            app._cleanup()