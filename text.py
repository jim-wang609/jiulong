class SegmentTree:
    # SegmentTree 类的构造函数
    def __init__(self, a: list[int]):
        n = len(a)  # 获取输入列表的长度
        # 计算线段树数组的大小，确保足够存储所有节点
        # (2 << (n - 1).bit_length()) 等同于 2 * (2**ceil(log2(n)))，即最近的2的幂的两倍
        self.max = [0] * (2 << (n - 1).bit_length())
        # 调用 build 方法初始化线段树
        self.build(a, 1, 0, n - 1)

    # 维护节点的最大值，即其左右子节点的最大值
    def maintain(self, o: int):
        self.max[o] = max(self.max[o * 2], self.max[o * 2 + 1])

    # 初始化线段树的递归函数
    def build(self, a: list[int], o: int, l: int, r: int):
        # 如果是叶子节点（区间只有一个元素）
        if l == r:
            self.max[o] = a[l]  # 将叶子节点的值设置为对应数组元素的值
            return
        m = (l + r) // 2  # 计算中间索引，将区间一分为二
        self.build(a, o * 2, l, m)  # 递归构建左子树
        self.build(a, o * 2 + 1, m + 1, r)  # 递归构建右子树
        self.maintain(o)  # 更新当前节点的最大值

    # 查找区间内第一个大于等于 x 的数，并将其更新为 -1，返回其索引（如果不存在则返回 -1）
    def find_first_and_update(self, o: int, l: int, r: int, x: int) -> int:
        # 如果当前节点的最大值小于 x，说明当前区间没有大于等于 x 的数
        if self.max[o] < x:
            return -1
        # 如果是叶子节点
        if l == r:
            self.max[o] = -1  # 将该叶子节点的值更新为 -1
            return l  # 返回该叶子节点的索引
        m = (l + r) // 2  # 计算中间索引
        i = self.find_first_and_update(o * 2, l, m, x)  # 首先尝试在左子树中查找
        # 如果左子树没有找到，则在右子树中查找
        if i < 0:
            i = self.find_first_and_update(o * 2 + 1, m + 1, r, x)
        self.maintain(o)  # 更新当前节点的最大值
        return i  # 返回找到的索引


class Solution:
    # 解决未放置水果数量问题的方法
    def numOfUnplacedFruits(self, fruits: list[int], baskets: list[int]) -> int:
        t = SegmentTree(baskets)  # 使用篮子容量列表初始化线段树
        n = len(baskets)  # 获取篮子数量
        ans = 0  # 初始化未放置水果的数量
        # 遍历每个水果
        for x in fruits:
            # 尝试找到一个能放置当前水果的篮子，并更新线段树
            if t.find_first_and_update(1, 0, n - 1, x) < 0:
                ans += 1  # 如果没有找到合适的篮子，则未放置水果数量加一
        return ans  # 返回最终未放置水果的数量