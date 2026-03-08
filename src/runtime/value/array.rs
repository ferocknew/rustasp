//! VBScript 数组实现
//!
//! 支持多维数组，使用扁平存储

use crate::runtime::{RuntimeError, Value};
use std::fmt;

/// VBScript 数组（支持多维）
///
/// # 多维数组表示
///
/// ```text
/// Dim a(2,3)  ' 3行4列的二维数组
///
/// dims = [3, 4]
/// data.len() = 12
///
/// 访问 a(1,2) 的扁平索引：
/// flat = 1 * 4 + 2 = 6
/// ```
#[derive(Debug, Clone)]
pub struct VbsArray {
    /// 每个维度的大小
    ///
    /// 例如 `Dim a(2,3)` → `dims = [3, 4]`
    /// - 第 0 维大小为 3（索引 0-2）
    /// - 第 1 维大小为 4（索引 0-3）
    pub dims: Vec<usize>,

    /// 扁平化的数据存储
    ///
    /// 使用 row-major 顺序存储
    pub data: Vec<Value>,
}

impl VbsArray {
    /// 创建新数组
    pub fn new(dims: Vec<usize>) -> Self {
        let size = if dims.is_empty() {
            0
        } else {
            dims.iter().product()
        };

        Self {
            dims,
            data: vec![Value::Empty; size],
        }
    }

    /// 创建一维数组（兼容 Array 函数）
    pub fn from_vec(data: Vec<Value>) -> Self {
        let len = data.len();
        Self {
            dims: vec![len],
            data,
        }
    }

    /// 计算多维索引的扁平索引（row-major）
    ///
    /// # 参数
    /// - `indices`: 多维索引，例如 `[1, 2]` 表示 `a(1,2)`
    ///
    /// # 返回
    /// - `Some(usize)`: 扁平索引
    /// - `None`: 索引越界或维度不匹配
    ///
    /// # 示例
    /// ```text
    /// dims = [3, 4]  // 3行4列
    ///
    /// a(0,0) → flat = 0 * 4 + 0 = 0
    /// a(0,1) → flat = 0 * 4 + 1 = 1
    /// a(1,0) → flat = 1 * 4 + 0 = 4
    /// a(1,2) → flat = 1 * 4 + 2 = 6
    /// a(2,3) → flat = 2 * 4 + 3 = 11
    /// ```
    pub fn flat_index(&self, indices: &[usize]) -> Option<usize> {
        // 维度检查
        if indices.len() != self.dims.len() {
            return None;
        }

        // 计算扁平索引
        let mut stride = 1;
        let mut flat = 0;

        for i in (0..self.dims.len()).rev() {
            // 边界检查
            if indices[i] >= self.dims[i] {
                return None;
            }

            flat += indices[i] * stride;
            stride *= self.dims[i];
        }

        Some(flat)
    }

    /// 获取数组总大小
    pub fn len(&self) -> usize {
        self.data.len()
    }

    /// 检查数组是否为空
    pub fn is_empty(&self) -> bool {
        self.data.is_empty()
    }

    /// 获取维度数量
    pub fn rank(&self) -> usize {
        self.dims.len()
    }

    /// 获取指定维度的大小
    pub fn dim_size(&self, dim: usize) -> Option<usize> {
        self.dims.get(dim).copied()
    }

    /// 调整数组大小（用于 ReDim）
    pub fn redim(&mut self, new_dims: Vec<usize>, preserve: bool) {
        let new_size = if new_dims.is_empty() {
            0
        } else {
            new_dims.iter().product()
        };

        if preserve {
            // 保留原有数据
            let old_data = std::mem::replace(&mut self.data, vec![Value::Empty; new_size]);
            let copy_len = old_data.len().min(new_size);

            for i in 0..copy_len {
                self.data[i] = old_data[i].clone();
            }
        } else {
            // 不保留，直接清空
            self.data = vec![Value::Empty; new_size];
        }

        self.dims = new_dims;
    }
}

impl fmt::Display for VbsArray {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if self.rank() == 1 {
            // 一维数组：普通显示
            let items: Vec<String> = self.data.iter().map(|v| v.to_string()).collect();
            write!(f, "[{}]", items.join(", "))
        } else {
            // 多维数组：显示维度信息
            write!(f, "[{:?} array, {} elements]", self.dims, self.data.len())
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_flat_index_1d() {
        let arr = VbsArray::new(vec![5]);
        assert_eq!(arr.flat_index(&[0]), Some(0));
        assert_eq!(arr.flat_index(&[4]), Some(4));
        assert_eq!(arr.flat_index(&[5]), None);
    }

    #[test]
    fn test_flat_index_2d() {
        // 3x4 数组
        let arr = VbsArray::new(vec![3, 4]);

        // row-major: flat = row * cols + col
        assert_eq!(arr.flat_index(&[0, 0]), Some(0));
        assert_eq!(arr.flat_index(&[0, 1]), Some(1));
        assert_eq!(arr.flat_index(&[0, 3]), Some(3));
        assert_eq!(arr.flat_index(&[1, 0]), Some(4));
        assert_eq!(arr.flat_index(&[1, 2]), Some(6));
        assert_eq!(arr.flat_index(&[2, 3]), Some(11));

        // 越界
        assert_eq!(arr.flat_index(&[3, 0]), None);
        assert_eq!(arr.flat_index(&[0, 4]), None);
        assert_eq!(arr.flat_index(&[1]), None); // 维度不匹配
    }

    #[test]
    fn test_flat_index_3d() {
        // 2x3x4 数组
        let arr = VbsArray::new(vec![2, 3, 4]);

        // flat = i * (3*4) + j * 4 + k
        assert_eq!(arr.flat_index(&[0, 0, 0]), Some(0));
        assert_eq!(arr.flat_index(&[0, 0, 1]), Some(1));
        assert_eq!(arr.flat_index(&[0, 1, 0]), Some(4));
        assert_eq!(arr.flat_index(&[1, 0, 0]), Some(12));
        assert_eq!(arr.flat_index(&[1, 2, 3]), Some(23));
    }

    #[test]
    fn test_redim() {
        let mut arr = VbsArray::new(vec![3]);
        arr.data[0] = Value::Number(1.0);
        arr.data[1] = Value::Number(2.0);
        arr.data[2] = Value::Number(3.0);

        // ReDim Preserve 扩展
        arr.redim(vec![5], true);
        assert_eq!(arr.dims, vec![5]);
        assert_eq!(arr.data.len(), 5);
        assert_eq!(arr.data[0], Value::Number(1.0));
        assert_eq!(arr.data[1], Value::Number(2.0));
        assert_eq!(arr.data[2], Value::Number(3.0));
        assert_eq!(arr.data[3], Value::Empty);

        // ReDim 不保留
        arr.redim(vec![2], false);
        assert_eq!(arr.dims, vec![2]);
        assert_eq!(arr.data.len(), 2);
        assert_eq!(arr.data[0], Value::Empty);
    }
}
