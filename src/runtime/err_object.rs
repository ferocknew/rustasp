//! Err 对象

/// VBScript Err 对象
#[derive(Debug, Clone)]
pub struct ErrObject {
    /// 错误号
    pub number: i32,
    /// 错误描述
    pub description: String,
}

impl ErrObject {
    /// 创建新的 Err 对象
    pub fn new() -> Self {
        ErrObject {
            number: 0,
            description: String::new(),
        }
    }

    /// 设置错误信息
    pub fn set(&mut self, number: i32, description: String) {
        self.number = number;
        self.description = description;
    }

    /// 清除错误信息
    pub fn clear(&mut self) {
        self.number = 0;
        self.description.clear();
    }

    /// 获取错误号
    pub fn get_number(&self) -> i32 {
        self.number
    }

    /// 获取错误描述
    pub fn get_description(&self) -> &str {
        &self.description
    }
}

impl Default for ErrObject {
    fn default() -> Self {
        Self::new()
    }
}

/// VBScript 错误代码常量
pub mod vb_error {
    /// 除零错误
    pub const DIVISION_BY_ZERO: i32 = 11;
    /// 类型不匹配
    pub const TYPE_MISMATCH: i32 = 13;
    /// 对象必需
    pub const OBJECT_REQUIRED: i32 = 424;
    /// 未定义的函数
    pub const UNDEFINED_FUNCTION: i32 = 1001;
    /// 数组下标越界
    pub const SUBSCRIPT_OUT_OF_RANGE: i32 = 9;
}
