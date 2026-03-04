//! 值类型模块
//!
//! VBScript 弱类型系统的实现

mod value;
mod conversion;
mod operators;
mod compare;
mod display;

pub use value::Value;
pub use conversion::ValueConversion;
pub use operators::ValueOps;
pub use compare::ValueCompare;
