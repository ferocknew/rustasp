//! 值类型模块
//!
//! VBScript 弱类型系统的实现

mod compare;
mod conversion;
mod display;
mod operators;
mod value;

pub use compare::ValueCompare;
pub use conversion::ValueConversion;
pub use operators::ValueOps;
pub use value::Value;
