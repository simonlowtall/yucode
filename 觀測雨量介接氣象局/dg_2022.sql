-- --------------------------------------------------------
-- 主機:                           211.23.167.206
-- 伺服器版本:                        5.5.29 - MySQL Community Server (GPL)
-- 伺服器作業系統:                      Win32
-- HeidiSQL 版本:                  10.3.0.5771
-- --------------------------------------------------------

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET NAMES utf8 */;
/*!50503 SET NAMES utf8mb4 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;


-- 傾印 dg_2022 的資料庫結構
CREATE DATABASE IF NOT EXISTS `dg_2022` /*!40100 DEFAULT CHARACTER SET utf8 */;
USE `dg_2022`;

-- 傾印  資料表 dg_2022.rt_rainfall 結構
CREATE TABLE IF NOT EXISTS `rt_rainfall` (
  `Rain_Station_ST_NO` varchar(25) NOT NULL,
  `DateTime` datetime NOT NULL,
  `Rainfall` double DEFAULT NULL,
  PRIMARY KEY (`Rain_Station_ST_NO`,`DateTime`),
  KEY `idx_rr_stn_dt` (`Rain_Station_ST_NO`,`DateTime`),
  KEY `idx_rr_datetime` (`DateTime`),
  KEY `Index_rt_rainfall` (`Rain_Station_ST_NO`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 取消選取資料匯出。

/*!40101 SET SQL_MODE=IFNULL(@OLD_SQL_MODE, '') */;
/*!40014 SET FOREIGN_KEY_CHECKS=IF(@OLD_FOREIGN_KEY_CHECKS IS NULL, 1, @OLD_FOREIGN_KEY_CHECKS) */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
