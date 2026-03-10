//! ASP 内置对象模块
//!
//! 包含所有 ASP 内置对象的实现：Request, Response, Session, Server

mod dictionary;
mod factory;
mod filesystemobject;
mod request;
mod response;
mod server;
mod session;
mod session_manager;
mod session_store;
mod xmlhttp;

// 重导出 Dictionary 对象
pub use dictionary::Dictionary;
// 重导出对象工厂
pub use factory::{create_object, get_supported_objects, is_whitelisted};
// 重导出 FileSystemObject 对象
pub use filesystemobject::FileSystemObject;
// 重导出 Request 对象
pub use request::Request;
// 重导出 Response 对象
pub use response::Response;
// 重导出 Session 对象
pub use session::{Session, SessionData};
// 重导出 Server 对象
pub use server::Server;
// 重导出 SessionManager
pub use session_manager::SessionManager;
// 重导出 Session 存储相关
pub use session_store::{create_store, JsonFileStore, MemoryStore, SessionStore};
// 重导出 XmlHttp 对象
pub use xmlhttp::XmlHttp;
