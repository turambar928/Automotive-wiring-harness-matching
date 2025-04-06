import win32com.client
import pythoncom
import time
import math
import random
import numpy as np


class CATIAV5Controller:
    def __init__(self):
        pythoncom.CoInitialize()
        try:
            self.catia = win32com.client.Dispatch("CATIA.Application")
            self.catia.Visible = True
            print("=== CATIA V5 自动化绘图 ===")
            print("CATIA 连接成功")

            # 定义四个矩形的固定长宽和初始位置
            self.rectangles = [
                {'width': 15, 'height': 8, 'x': 110, 'y': 110, 'angle': 0},
                {'width': 12, 'height': 6, 'x': 85, 'y': 105, 'angle': math.pi / 4},
                {'width': 10, 'height': 10, 'x': 100, 'y': 80, 'angle': math.pi / 6},
                {'width': 8, 'height': 12, 'x': 115, 'y': 90, 'angle': -math.pi / 4}
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

    def draw_circle(self, sketch, x=0, y=0, radius=50):
        """绘制圆形"""
        try:
            factory = sketch.OpenEdition()
            circle = factory.CreateClosedCircle(x, y, radius)
            sketch.CloseEdition()
            self.part.Update()
            print(f"圆形创建成功 (中心: ({x}, {y}), 半径: {radius})")
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

    def get_rectangle_corners(self, rect):
        """获取矩形的四个角点坐标"""
        x, y = rect['x'], rect['y']
        half_w = rect['width'] / 2
        half_h = rect['height'] / 2
        angle = rect['angle']

        cos_a = math.cos(angle)
        sin_a = math.sin(angle)

        corners = []
        for dx, dy in [(-half_w, -half_h), (half_w, -half_h),
                       (half_w, half_h), (-half_w, half_h)]:
            rot_x = x + dx * cos_a - dy * sin_a
            rot_y = y + dx * sin_a + dy * cos_a
            corners.append((rot_x, rot_y))

        return corners

    def check_overlap(self, rect1, rect2):
        """检查两个矩形是否重叠(分离轴定理)"""

        def project(poly, axis):
            dots = [p[0] * axis[0] + p[1] * axis[1] for p in poly]
            return min(dots), max(dots)

        corners1 = self.get_rectangle_corners(rect1)
        corners2 = self.get_rectangle_corners(rect2)

        # 检查矩形1的边
        edges = []
        for i in range(4):
            x1, y1 = corners1[i]
            x2, y2 = corners1[(i + 1) % 4]
            edge = (x2 - x1, y2 - y1)
            length = math.sqrt(edge[0] ** 2 + edge[1] ** 2)
            if length > 0:
                normal = (-edge[1] / length, edge[0] / length)
                edges.append(normal)

        # 检查矩形2的边
        for i in range(4):
            x1, y1 = corners2[i]
            x2, y2 = corners2[(i + 1) % 4]
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

    def evaluate_solution(self, solution):
        """评估解决方案的质量"""
        circle_x, circle_y, circle_r = solution['circle']
        rectangles = solution['rectangles']

        # 检查矩形间是否重叠
        overlap_penalty = 0
        for i in range(len(rectangles)):
            for j in range(i + 1, len(rectangles)):
                if self.check_overlap(rectangles[i], rectangles[j]):
                    overlap_penalty += 1000  # 大惩罚项

        # 检查所有矩形是否在圆内
        max_distance = 0
        for rect in rectangles:
            corners = self.get_rectangle_corners(rect)
            for corner_x, corner_y in corners:
                distance = math.sqrt((corner_x - circle_x) ** 2 + (corner_y - circle_y) ** 2)
                if distance > max_distance:
                    max_distance = distance

        # 适应度：圆半径越小越好 + 重叠惩罚
        if max_distance <= circle_r:
            return circle_r + overlap_penalty  # 越小越好
        else:
            return float('inf')  # 无效解

    def simulated_annealing(self, iterations=1000, temp=1000, cooling_rate=0.99):
        """模拟退火优化矩形位置和包围圆"""
        # 初始化当前解
        current_solution = {
            'circle': [100, 100, 50],  # 初始圆参数[x, y, radius]
            'rectangles': [r.copy() for r in self.rectangles]  # 复制初始矩形
        }

        # 初始化最佳解
        best_solution = {
            'circle': current_solution['circle'].copy(),
            'rectangles': [r.copy() for r in current_solution['rectangles']]
        }

        for i in range(iterations):
            # 生成新解 - 深拷贝当前解
            new_solution = {
                'circle': current_solution['circle'].copy(),
                'rectangles': [r.copy() for r in current_solution['rectangles']]
            }

            # 随机调整每个矩形
            for rect in new_solution['rectangles']:
                # 随机平移
                rect['x'] += random.uniform(-5, 5)
                rect['y'] += random.uniform(-5, 5)
                # 随机旋转
                rect['angle'] += random.uniform(-0.2, 0.2)

            # 调整圆参数
            new_solution['circle'][0] += random.uniform(-3, 3)  # x
            new_solution['circle'][1] += random.uniform(-3, 3)  # y
            new_solution['circle'][2] = max(10, new_solution['circle'][2] + random.uniform(-2, 2))  # 半径

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

                if new_fitness < self.evaluate_solution(best_solution):
                    best_solution = {
                        'circle': new_solution['circle'].copy(),
                        'rectangles': [r.copy() for r in new_solution['rectangles']]
                    }

            # 降低温度
            temp *= cooling_rate

            # 打印进度
            if i % 100 == 0:
                print(f"Iteration {i}: Temp={temp:.2f}, Best Radius={best_solution['circle'][2]:.2f}")

        return best_solution

    def draw_solution(self, solution):
        """绘制解决方案"""
        try:
            # 绘制矩形
            for i, rect in enumerate(solution['rectangles']):
                if not self.draw_rectangle(self.sketch, rect['x'], rect['y'],
                                           rect['width'], rect['height'], rect['angle']):
                    raise Exception(f"无法绘制第 {i + 1} 个矩形")

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
            sketches = self.main_body.Sketches
            sketch_count = sketches.Count

            # 删除当前 sketch（如果存在）
            if self.sketch:
                for i in range(1, sketch_count + 1):
                    item = sketches.Item(i)
                    if item.Name == self.sketch.Name:
                        item.Delete()
                        print(f"已删除旧草图: {self.sketch.Name}")
                        break

            # 创建新的草图
            self.sketch = self.create_sketch(plane=plane)
            return True
        except Exception as e:
            print(f"重置草图失败: {e}")
            return False

    def run(self):
        """运行主程序"""
        try:
            if not self.create_part():
                return False

            # 1. 绘制初始布局
            self.sketch = self.create_sketch()
            print("\n绘制初始布局...")
            for rect in self.rectangles:
                self.draw_rectangle(self.sketch, rect['x'], rect['y'],
                                    rect['width'], rect['height'], rect['angle'])

            # 2. 优化布局
            input("\n按回车键开始优化布局...")
            print("\n正在优化布局(可能需要几分钟)...")
            self.best_solution = self.simulated_annealing(iterations=2000)

            # 3. 清除初始布局并绘制优化结果
            self.reset_sketch()

            print("\n绘制优化后的布局...")
            if not self.draw_solution(self.best_solution):
                return False

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
        finally:
            self._cleanup()


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