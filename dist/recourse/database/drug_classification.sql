/*
 Navicat Premium Data Transfer

 Source Server         : care_mysql
 Source Server Type    : MySQL
 Source Server Version : 80027
 Source Host           : localhost:3306
 Source Schema         : drug_classification

 Target Server Type    : MySQL
 Target Server Version : 80027
 File Encoding         : 65001

 Date: 11/01/2023 17:14:11
*/

SET NAMES utf8mb4;
SET FOREIGN_KEY_CHECKS = 0;

-- ----------------------------
-- Table structure for 4+7_base_drug
-- ----------------------------
DROP TABLE IF EXISTS `4+7_base_drug`;
CREATE TABLE `4+7_base_drug`  (
  `序号` int(0) NOT NULL,
  `药品名称` varchar(255) CHARACTER SET utf8 COLLATE utf8_bin NULL DEFAULT NULL,
  `规格` varchar(255) CHARACTER SET utf8 COLLATE utf8_bin NULL DEFAULT NULL,
  `单位` varchar(255) CHARACTER SET utf8 COLLATE utf8_bin NULL DEFAULT NULL,
  `单价` varchar(255) CHARACTER SET utf8 COLLATE utf8_bin NULL DEFAULT NULL,
  `生产厂家` varchar(255) CHARACTER SET utf8 COLLATE utf8_bin NULL DEFAULT NULL,
  PRIMARY KEY (`序号`) USING BTREE
) ENGINE = InnoDB CHARACTER SET = utf8 COLLATE = utf8_bin ROW_FORMAT = Dynamic;

-- ----------------------------
-- Records of 4+7_base_drug
-- ----------------------------
INSERT INTO `4+7_base_drug` VALUES (1, '阿莫西林颗粒', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (2, '阿莫西林胶囊', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (3, '阿奇霉素胶囊', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (4, '阿托伐他汀钙片', '10mg*14片', '盒', '2.8', NULL);
INSERT INTO `4+7_base_drug` VALUES (5, '阿托伐他汀钙片', '20mg*28片', '盒', '9.52', NULL);
INSERT INTO `4+7_base_drug` VALUES (6, '奥氮平', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (7, '苯磺酸氨氯地平片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (8, '布洛芬缓释胶囊', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (9, '布洛芬颗粒', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (10, '多潘立酮片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (11, '恩替卡韦平', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (12, '非那雄胺片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (13, '复方氨基酸注射液', NULL, '瓶', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (14, '格列吡嗪片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (15, '格列美脲片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (16, '甲硝唑片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (17, '胶体果胶铋胶囊', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (18, '酒石酸美托洛尔片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (19, '卡托普利片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (20, '克拉霉素片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (21, '利伐沙班片限制药', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (22, '硫酸氢氯吡格雷片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (23, '氯雷他定片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (24, '马来酸依那普利片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (25, '蒙脱石散', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (26, '尼莫地平片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (27, '诺氟沙星片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (28, '瑞格列奈片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (29, '瑞舒伐他汀钙片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (30, '替格瑞洛片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (31, '头孢呋辛酯片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (32, '维生素B6片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (33, '吸入用布地奈德混悬液', NULL, '支', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (34, '吸入用硫酸沙丁胺醇溶液', NULL, '支', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (35, '缬沙坦胶囊', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (36, '辛伐他汀片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (37, '盐酸氨溴索分散片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (38, '盐酸二甲双胍缓释片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (39, '盐酸二甲双胍片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (40, '盐酸坦索罗辛缓释胶囊', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (41, '吲哒帕胺片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (42, '注射用甲泼尼龙琥珀酸钠', NULL, '支', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (43, '注射用头孢呋辛钠', NULL, '支', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (44, '左氧氟沙星片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (45, '苯磺酸氨氯地平片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (46, '注射用法莫替丁', NULL, '支', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (47, '盐酸雷尼替丁注射液', '10支/盒', '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (48, '盐酸甲氧氯普胺注射液', '10支/盒', '支', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (49, '金水宝片限制药', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (50, '活血止痛胶囊', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (51, '阿卡波糖片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (52, '枸橼酸莫沙必利片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (53, '碳酸氢钠注射液', NULL, '支', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (54, '血府逐瘀丸', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (55, '双黄连口服液', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (57, '甘露醇注射液', NULL, '瓶', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (58, '尼群地平片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (59, '肾衰宁', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (60, '血塞通', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (61, '银杏叶', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (62, '乳果糖口服溶液', NULL, '盒', '0', NULL);
INSERT INTO `4+7_base_drug` VALUES (63, '奥美拉唑肠溶胶囊', NULL, '盒', '0', NULL);

-- ----------------------------
-- Table structure for 4+7_non_base_drug
-- ----------------------------
DROP TABLE IF EXISTS `4+7_non_base_drug`;
CREATE TABLE `4+7_non_base_drug`  (
  `序号` int(0) NOT NULL,
  `药品名称` varchar(255) CHARACTER SET utf8 COLLATE utf8_bin NULL DEFAULT NULL,
  `规格` varchar(255) CHARACTER SET utf8 COLLATE utf8_bin NULL DEFAULT NULL,
  `单位` varchar(255) CHARACTER SET utf8 COLLATE utf8_bin NULL DEFAULT NULL,
  `单价` varchar(255) CHARACTER SET utf8 COLLATE utf8_bin NULL DEFAULT NULL,
  `生产厂家` varchar(255) CHARACTER SET utf8 COLLATE utf8_bin NULL DEFAULT NULL,
  PRIMARY KEY (`序号`) USING BTREE
) ENGINE = InnoDB CHARACTER SET = utf8 COLLATE = utf8_bin ROW_FORMAT = Dynamic;

-- ----------------------------
-- Records of 4+7_non_base_drug
-- ----------------------------
INSERT INTO `4+7_non_base_drug` VALUES (1, '阿德福韦酯片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (2, '玻璃酸钠注射液', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (3, '地氯雷他定片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (4, '厄贝沙坦片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (5, '厄贝沙坦氢氯噻嗪片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (6, '复方利血平片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (7, '格列齐特缓释片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (8, '甲钴胺片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (9, '坎地沙坦酯片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (10, '氯沙坦钾片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (11, '孟鲁司特钠片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (12, '泮托拉唑钠肠溶片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (13, '匹伐他汀钙片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (14, '塞来昔布胶囊', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (15, '天麻素注射液', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (16, '头孢克洛胶囊', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (17, '盐酸氨基葡萄糖胶囊', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (18, '盐酸氨溴索注射液', NULL, '支', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (19, '盐酸西替利嗪片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (20, '盐酸左西替利嗪片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (21, '注射液头孢曲松钠', NULL, '支', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (22, '注射用阿奇霉素', NULL, '支', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (23, '注射用美洛西林钠', NULL, '支', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (24, '注射用七叶皂苷钠', NULL, '支', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (25, '注射用头孢曲松钠', NULL, '支', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (26, '左氧氟沙星滴眼液', NULL, '支', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (27, '左氧氟沙星氯化钠注射液', NULL, '袋', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (28, '左乙拉西坦片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (29, '曲克芦丁片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (30, '马来酸氯苯那敏注射液.非', NULL, '支', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (31, '氯化钾注射液', NULL, '支', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (32, '脂肪乳注射用C14-24)', NULL, '瓶', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (33, '奥硝唑氯化钠注射液', NULL, '瓶', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (34, '格列齐特缓释片', NULL, '盒', '0', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (35, '康复新液', '每瓶装30ml*4瓶', '瓶', '13.79', '内蒙古京新药业有限公司');
INSERT INTO `4+7_non_base_drug` VALUES (36, '小儿布洛芬栓', '50mg8枚', '盒', '26.64', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (37, '右旋布洛芬口服混悬液', '100ml', '盒', '40', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (38, '右旋布洛芬栓', '50mg*12粒', '盒', '39.15', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (39, '抗病毒颗粒', '4g*12袋', '盒', '17.04', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (40, '四季抗病毒合剂', '10ml*6支', '盒', '49.5', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (41, '利巴韦林泡腾颗粒', '50mg*36袋', '盒', '38', NULL);
INSERT INTO `4+7_non_base_drug` VALUES (42, '金花清感颗粒', '6袋', '盒', '53.4', NULL);

-- ----------------------------
-- Table structure for base_drug
-- ----------------------------
DROP TABLE IF EXISTS `base_drug`;
CREATE TABLE `base_drug`  (
  `序号` int(0) NOT NULL,
  `药品名称` varchar(255) CHARACTER SET utf8 COLLATE utf8_bin NULL DEFAULT NULL,
  `规格` varchar(255) CHARACTER SET utf8 COLLATE utf8_bin NULL DEFAULT NULL,
  `单位` varchar(255) CHARACTER SET utf8 COLLATE utf8_bin NULL DEFAULT NULL,
  `单价` varchar(255) CHARACTER SET utf8 COLLATE utf8_bin NULL DEFAULT NULL,
  `生产厂家` varchar(255) CHARACTER SET utf8 COLLATE utf8_bin NULL DEFAULT NULL,
  PRIMARY KEY (`序号`) USING BTREE
) ENGINE = InnoDB CHARACTER SET = utf8 COLLATE = utf8_bin ROW_FORMAT = Dynamic;

-- ----------------------------
-- Records of base_drug
-- ----------------------------
INSERT INTO `base_drug` VALUES (1, '（10%）葡萄糖注射液', '500ml', '30/件', '47.1', '石家庄四药股份有限公司');
INSERT INTO `base_drug` VALUES (2, '（5%）葡萄糖注射液', '100ml', '90/件', '387', '石家庄四药股份有限公司');
INSERT INTO `base_drug` VALUES (3, '（5%）葡萄糖注射液', '250ml', '60/件', '69', '石家庄四药股份有限公司');
INSERT INTO `base_drug` VALUES (4, '（5%）葡萄糖注射液', '500ml', '30/件', '47.1', '石家庄四药股份有限公司');
INSERT INTO `base_drug` VALUES (5, '阿苯达唑片', '0.2g*10s', '盒', '8.83', '中美史克');
INSERT INTO `base_drug` VALUES (6, '阿莫西林克拉维酸钾颗粒', '0.2285*6袋', '盒', '23.33', '黑龙江全乐制药');
INSERT INTO `base_drug` VALUES (7, '阿莫西林克拉维酸钾片', '0.2285g*6片', '盒', '13.48', '四川');
INSERT INTO `base_drug` VALUES (8, '阿奇霉素颗粒', NULL, '盒', '0', NULL);
INSERT INTO `base_drug` VALUES (9, '阿司匹林肠溶片', '50mg*100p', '瓶', '7.6', '山西同达药业');
INSERT INTO `base_drug` VALUES (10, '阿司匹林肠溶片', '100mg36t', '盒', '17.84', '石药集团');
INSERT INTO `base_drug` VALUES (11, '氨茶碱片', '0.1克*100片', '瓶', '8.48', '天津力生制药股份有限公司');
INSERT INTO `base_drug` VALUES (12, '氨茶碱注射液', '0.25g*10支', '盒', '39', '石药银湖制药有限公司');
INSERT INTO `base_drug` VALUES (13, '柏子养心丸', '9g*10w', '盒', '6.25', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (14, '板蓝根颗粒', '10g*15d', '包', '11', '广州白云山制药股份有限公司广州白云山制药总厂');
INSERT INTO `base_drug` VALUES (15, '保妇康栓', '1.74g', '8粒/盒', '34.72', '海南碧凯药业有限公司');
INSERT INTO `base_drug` VALUES (16, '保和颗粒', '4.5g*10d', '盒', '21.7', '太极集团');
INSERT INTO `base_drug` VALUES (17, '保和丸', '9g*10丸', '盒', '5.2', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (18, '苯磺酸左氨氯地平片', '2.5mg*21s', '盒', '36', '贵州联盛药业');
INSERT INTO `base_drug` VALUES (19, '苯唑西林钠胶囊', '36粒', '盒', '39.33', '四川制药');
INSERT INTO `base_drug` VALUES (20, '鼻炎康片', '0.37g96片', '瓶', '20.58', '国药集团');
INSERT INTO `base_drug` VALUES (21, '参苓白术散', '9g*10袋', '盒', '0', '云南腾');
INSERT INTO `base_drug` VALUES (22, '参芪降糖胶囊', '0.35*90粒', '盒', '48.15', '河南羚锐制');
INSERT INTO `base_drug` VALUES (23, '茶碱缓释片', '0.1g*30t', '盒', '8.7', NULL);
INSERT INTO `base_drug` VALUES (24, '柴胡注射液', '2ml*10支', '盒', '3.9', '河南福森');
INSERT INTO `base_drug` VALUES (25, '醋酸泼尼松片', '5mg100p', '瓶', '6.2', '山东新华');
INSERT INTO `base_drug` VALUES (26, '达格列净片', '10mg*7片*2板', '盒', '61.04', '阿斯利康制药有限公司');
INSERT INTO `base_drug` VALUES (27, '丹珍头痛胶囊', '0.5g*24粒', '盒', '33.1', '青海益欣药业');
INSERT INTO `base_drug` VALUES (28, '单硝酸异山梨酯片', '20mg*48s', '盒', '37.97', '鲁南贝特');
INSERT INTO `base_drug` VALUES (29, '灯盏花素片', '20mg*18s*2板', '盒', '49.8', '逐成药业');
INSERT INTO `base_drug` VALUES (30, '地塞米松磷酸钠注射液', '5mg*1ml*10支', '盒', '4.5', '石药银湖制药有限公司');
INSERT INTO `base_drug` VALUES (31, '地衣芽孢杆菌活菌胶囊', '0.25g*20粒', '盒', '20.34', '东北制药集团公司沈阳第一制药厂');
INSERT INTO `base_drug` VALUES (32, '地衣芽孢杆菌活菌颗粒（整肠生）', '0.5g*12袋', '盒', '0', '东北制药集团沈阳第一制药有限公司');
INSERT INTO `base_drug` VALUES (33, '颠茄片', '10mg*100p', '瓶', '5.18', '天津天士力制药股份有限公司');
INSERT INTO `base_drug` VALUES (34, '丁酸氢化可的松乳膏', '20g', '支', '14.6', '天津太平洋制药有限公司');
INSERT INTO `base_drug` VALUES (35, '对乙酰氨基酚口服液混悬液', '100ml*3.2g', '瓶', '12.8', '上海强生');
INSERT INTO `base_drug` VALUES (36, '对乙酰氨基酚干混悬剂', '1.01g*10袋', '盒', '17.75', '青岛黄海');
INSERT INTO `base_drug` VALUES (37, '防风通圣丸', '6g10袋', '盒', '6.5', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (38, '非诺贝特分散片', '0.1g*24片', '盒', '30.42', '东莞市金美药业有限公司');
INSERT INTO `base_drug` VALUES (39, '风湿骨痛片', '0.36g*48片', '盒', '30.11', '安徽美欣制药有限公司');
INSERT INTO `base_drug` VALUES (40, '风湿液', '250ml', '盒', '46.68', '天津达仁堂京万红药业');
INSERT INTO `base_drug` VALUES (41, '呋喃妥因片', '50mg100片', '瓶', '7.7', '天津力生制药股份有限公司');
INSERT INTO `base_drug` VALUES (42, '呋塞米片', '20mg*100p', '瓶', '29.8', '天津力生制药股份有限公司');
INSERT INTO `base_drug` VALUES (43, '妇科千金片', '无', '144片/瓶', '34', '株洲千金药业股份有限公司');
INSERT INTO `base_drug` VALUES (44, '妇炎消胶囊', '0.45g*48粒', '盒', '30.3', '贵州益佰女子大药厂有限责任公司');
INSERT INTO `base_drug` VALUES (45, '附子理中丸', '9g*10丸', '盒', '5.83', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (46, '复方丹参滴丸', '27mg*180丸', '瓶', '26.08', '天津天士力制药股份有限公司');
INSERT INTO `base_drug` VALUES (47, '复方丹参片', '100片', '瓶', '26.13', '吉林道君');
INSERT INTO `base_drug` VALUES (48, '复方甘草片', '100片', '瓶', '9', '东北制药集团公司沈阳第一制药厂');
INSERT INTO `base_drug` VALUES (49, '复方黄柏液涂剂', '150ml', '1瓶/盒', '57.2', '山东汉方制药有限公司');
INSERT INTO `base_drug` VALUES (50, '感冒清热胶囊', '0.45*36粒', '盒', '27.34', '吉林敖东');
INSERT INTO `base_drug` VALUES (51, '感冒清热颗粒', '12g*6袋', '合', '15', '天圣制药');
INSERT INTO `base_drug` VALUES (52, '宫炎平胶囊', '0.2g*60粒', '盒', '23.5', '江西桔王药业');
INSERT INTO `base_drug` VALUES (53, '枸橼酸铋钾颗粒', '1g*56袋', '盒', '30.24', '国药集团');
INSERT INTO `base_drug` VALUES (54, '冠心苏合胶囊', '0.35g*20粒', '盒', '24.95', '吉林朝药');
INSERT INTO `base_drug` VALUES (55, '归脾丸', '9g*10丸', '盒', '6.15', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (56, '桂龙咳喘宁片', '0.41g*60片', '盒', '18.4', '江西药都仁和制药有限公司');
INSERT INTO `base_drug` VALUES (57, '桂枝茯苓丸', '6g*10丸', '盒', '6.1', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (58, '寒喘祖帕颗粒', '12g*6袋', '盒', '22.79', '新疆维吾药业');
INSERT INTO `base_drug` VALUES (59, '红金消结胶囊', '0.4g', '36粒/盒', '31.61', '云南佑生药业有限责任公司');
INSERT INTO `base_drug` VALUES (60, '红霉素软膏', '15g', '支', '1.8', '福元药业');
INSERT INTO `base_drug` VALUES (61, '琥珀酸亚铁片', '0.1g*18片', '盒', '29.54', '湖南九典制药');
INSERT INTO `base_drug` VALUES (62, '护肝片', '0.36g*100p', '瓶', '20.49', '上海宝龙安庆药业有限');
INSERT INTO `base_drug` VALUES (63, '花红颗粒', '每袋装2.5g（无蔗糖）', '12袋/盒', '22.49', '广西壮族自治区花红药业股份有限公司');
INSERT INTO `base_drug` VALUES (64, '华佗再造丸', '8g12袋', '盒', '68', '广州白云山制药股份有限公司广州白云山制药总厂');
INSERT INTO `base_drug` VALUES (65, '滑膜炎颗粒', '6g*6d', '盒', '60', '神威药业有限公司');
INSERT INTO `base_drug` VALUES (66, '槐角丸', '9g10丸', '盒', '5.93', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (67, '黄体酮注射液', '20mg', '盒', '3.6', '浙江仙居制药有限公司');
INSERT INTO `base_drug` VALUES (68, '活心丸', '20mg*30片', '盒', '46.02', '广州悦康生物制药有限公司');
INSERT INTO `base_drug` VALUES (69, '藿胆丸', '36g', '瓶', '5.8', '武汉太福');
INSERT INTO `base_drug` VALUES (70, '藿香正气水', '10ml*10支', '盒', '6', '四川泰华堂');
INSERT INTO `base_drug` VALUES (71, '藿香正气软胶囊', '0.45*24粒', '盒', '17.44', '神威药业有限公司');
INSERT INTO `base_drug` VALUES (72, '急支糖浆', '200ml', '瓶', '45', '太极集团重庆涪陵制药厂有限公司');
INSERT INTO `base_drug` VALUES (73, '加味左金丸', '6*12袋', '盒', '24.53', '江西药都樟树药业');
INSERT INTO `base_drug` VALUES (74, '甲硝唑氯化钠注射液', '100ml*0.5g', '瓶', '2.97', '石家庄四药股份有限公司');
INSERT INTO `base_drug` VALUES (75, '甲硝唑栓', '0.5g*10粒', '盒', '5.5', '马应龙药业');
INSERT INTO `base_drug` VALUES (76, '健儿消食口服液', '10ml*10z', '盒', '14.35', '回音必集团安徽制药有限公司');
INSERT INTO `base_drug` VALUES (77, '健脾生血颗粒', '5g*36袋', '盒', '34.02', '健民药业');
INSERT INTO `base_drug` VALUES (78, '健脾生血片', '每片重0.6克', '36片/盒', '24.27', '健民药业集团股份有限公司');
INSERT INTO `base_drug` VALUES (79, '接骨七厘丸', '每袋装1.5克', '盒', '20.93', '洛阳顺势药业有限公司');
INSERT INTO `base_drug` VALUES (80, '金匮肾气丸', '60g', '瓶', '21.99', '药都制药');
INSERT INTO `base_drug` VALUES (81, '金芪降糖片', '36片', '盒', '53.64', '天津中新药业集团有限公司');
INSERT INTO `base_drug` VALUES (82, '金钱胆通颗粒', '8g*10袋', '盒', '58.75', '贵州威门药业');
INSERT INTO `base_drug` VALUES (83, '金叶败毒颗粒', '每袋装10g*8袋', '盒', '49.8', '国药集团中联药业有限公司');
INSERT INTO `base_drug` VALUES (84, '京万红软膏', '每支装50g', '1支/盒', '67.5', '天津达仁堂京万红药业有限公司');
INSERT INTO `base_drug` VALUES (85, '颈复康颗粒', '5g*10d', '盒', '20.34', '承德颈复康药业集团有限公司');
INSERT INTO `base_drug` VALUES (86, '九味羌活颗粒', '5g*20袋', '盒', '28.8', '九寨沟天然药业');
INSERT INTO `base_drug` VALUES (87, '橘红颗粒', '11g*12袋', '盒', '21.4', '云南海沣');
INSERT INTO `base_drug` VALUES (88, '橘红丸', '6g*10丸', '盒', '5.1', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (89, '开塞露', '20ml', '支', '1.3', '上海运佳');
INSERT INTO `base_drug` VALUES (90, '口服补液盐III', '5.125g*6袋', '盒', '31.29', '西安安健有限公司');
INSERT INTO `base_drug` VALUES (91, '口炎清颗粒', '10g*12袋', '盒', '25.34', '广州白云山制药股份有限公司广州白云山制药总厂');
INSERT INTO `base_drug` VALUES (92, '坤泰胶囊', '每粒装0.5g', '36粒/盒', '26.89', '贵阳新天药业股份有限公司');
INSERT INTO `base_drug` VALUES (93, '来氟米特片', '10mg*12片', '盒', '28.7', '美罗药业有限公司');
INSERT INTO `base_drug` VALUES (94, '连花清瘟胶囊', '0.35g24', '盒', '11.09', '石家庄以岭药业股份有限公司');
INSERT INTO `base_drug` VALUES (95, '连花清瘟颗粒', '每袋装6g', '盒', '23.57', '北京以岭药业有限公司');
INSERT INTO `base_drug` VALUES (96, '硫酸庆大霉素注射液', '2ml8万单位', '支', '0.6', '河南润弘');
INSERT INTO `base_drug` VALUES (97, '六味地黄丸（蜜丸）', '9g*10丸', '盒', '6.15', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (98, '六味地黄丸（水丸）', '180丸', '盒', '21.8', '上海宝龙');
INSERT INTO `base_drug` VALUES (99, '癃清胶囊', '0.6g*48片', '盒', '39.64', '贵州远程制药有');
INSERT INTO `base_drug` VALUES (100, '螺内酯片', '20mg*100片', '瓶', '13.9525', '贵州远程制药');
INSERT INTO `base_drug` VALUES (101, '氯化钠注射液', '100ml', '80袋/件', '4.3', '石家庄四药股份有限公司');
INSERT INTO `base_drug` VALUES (102, '氯化钠注射液', '250ml', '60袋/件', '69', '石家庄四药股份有限公司');
INSERT INTO `base_drug` VALUES (103, '氯化钠注射液（塑瓶）', '500ml', '30/件', '44.41', '石家庄四药股份有限公司');
INSERT INTO `base_drug` VALUES (104, '氯化钠注射液', '500ml', '30/件', '59.4', '石家庄四药股份有限公司');
INSERT INTO `base_drug` VALUES (105, '麻仁润肠丸', '6g*10丸', '盒', '6.7', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (106, '马来酸依那普利叶酸片', '10mg*7t', '盒', '33.53', '深圳奥萨制药');
INSERT INTO `base_drug` VALUES (107, '马应龙麝香痔疮膏', '24g', '盒', '20.68', '马应龙药业');
INSERT INTO `base_drug` VALUES (108, '明目地黄丸', '9g*10丸', '盒', '5.51', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (109, '莫匹罗星软膏', '10g', '支', '20.18', NULL);
INSERT INTO `base_drug` VALUES (110, '脑安滴丸', '50mg', '盒', '25.91', '安徽意隆亚东');
INSERT INTO `base_drug` VALUES (111, '脑心通胶囊', '0.4g*36粒', '盒', '26', '陕西步长制药有限公司');
INSERT INTO `base_drug` VALUES (112, '牛黄解毒胶囊', '0.3g*12粒*4板', '盒', '31.99', '贵州济生药业有限公司');
INSERT INTO `base_drug` VALUES (113, '牛黄解毒丸', '3克*10丸', '盒', '4.66', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (114, '牛黄清感胶囊', '每粒装0.3g', '36粒/盒', '23.76', '黑龙江澳利达奈德制药有限公司');
INSERT INTO `base_drug` VALUES (115, '牛黄上清胶囊', '0.3g*36粒', '盒', '21.36', '江西康恩贝');
INSERT INTO `base_drug` VALUES (116, '牛黄上清丸', '6克*10丸', '盒', '5.1', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (117, '排石颗粒', '5g*12袋', '盒', '29.2', '南京同仁堂药业有限公司');
INSERT INTO `base_drug` VALUES (118, '匹维溴铵片', '50mg*20片', '盒', '29.37', '北京福元医药');
INSERT INTO `base_drug` VALUES (119, '葡萄糖酸钙片', '0.5g*100t', '瓶', '19.9', '天津力生制药股份有限公司');
INSERT INTO `base_drug` VALUES (120, '葡萄糖酸钙注射液', '1g', '支', '9.6', NULL);
INSERT INTO `base_drug` VALUES (121, '葡萄糖注射液', '20ml', '支', '1.13', NULL);
INSERT INTO `base_drug` VALUES (122, '普乐安片', NULL, '瓶', '0', NULL);
INSERT INTO `base_drug` VALUES (123, '芪参益气滴丸', '0.5g*15袋', '盒', '33.29', '天士力医药集团');
INSERT INTO `base_drug` VALUES (124, '芪苈强心胶囊', '每粒装0.3g', '36粒/盒', '32.74', '石家庄以岭药业股份有限公司');
INSERT INTO `base_drug` VALUES (125, '杞菊地黄丸', '9克*10丸', '盒', '6.2', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (126, '气滞胃痛片', '0.5*30片', '盒', '45', '辽宁华润本溪三药有限公司');
INSERT INTO `base_drug` VALUES (127, '氢氯噻嗪片', '10mg*100s', '瓶', '4.4', '山西云鹏');
INSERT INTO `base_drug` VALUES (128, '清开灵颗粒', '3g*24袋', '盒', '30.36', '哈尔滨一洲制药');
INSERT INTO `base_drug` VALUES (129, '清热八味胶囊', '0.3g*12袋', '盒', '35', '内蒙古奥特奇蒙古药业');
INSERT INTO `base_drug` VALUES (130, '清宣止咳颗粒', '10g', '6袋/盒', '18.21', '江苏苏中药业集团股份有限公司');
INSERT INTO `base_drug` VALUES (131, '乳酶生片', '0.3g*60片', '盒', '29.45', '桂林南药股份');
INSERT INTO `base_drug` VALUES (132, '乳癖消片', '0.34g*100p', '瓶', '24.69', '辽宁好护士药业集团有限责任公司');
INSERT INTO `base_drug` VALUES (133, '乳酸钠林格注射液', '500ml', '瓶', '6.6', '石家庄四药股份有限公司');
INSERT INTO `base_drug` VALUES (134, '三金片', '72片', '盒', '29.2', '桂林三金');
INSERT INTO `base_drug` VALUES (135, '伤科接骨片', '0.33g*60P', '盒', '35.82', '大连美罗中药厂有限公司');
INSERT INTO `base_drug` VALUES (136, '蛇胆川贝液', '10ml', '盒', '5.8', '武汉太福制药有限公司');
INSERT INTO `base_drug` VALUES (137, '麝香保心丸', '22.5mg*42粒', '盒', '29.8', '上海和黄药业');
INSERT INTO `base_drug` VALUES (138, '麝香追风止痛膏', '7cm×10cm*4贴', '盒', '48', '重庆希尔安药业有限公司');
INSERT INTO `base_drug` VALUES (139, '生脉饮口服液', '10ml*8支', '盒', '10.93', '四川');
INSERT INTO `base_drug` VALUES (140, '生血宝颗粒', '4g*8袋', '盒', '27.81', '湖南康寿制药有限公司');
INSERT INTO `base_drug` VALUES (141, '舒筋活血片', '0.3g*48片', '盒', '12.66', '库尔勒龙');
INSERT INTO `base_drug` VALUES (142, '双氯芬酸钠缓释胶囊', '50mg*20片', '盒', '24.4', '南京制药');
INSERT INTO `base_drug` VALUES (143, '双歧杆菌三联活菌胶囊', '0.21g', '36粒/瓶', '30.22', '上海上药信谊药厂有限公司');
INSERT INTO `base_drug` VALUES (144, '四妙丸', '6g*12袋', '盒', '42.75', '吉林紫鑫药业');
INSERT INTO `base_drug` VALUES (145, '松龄血脉康胶囊', '每粒装0.5g', '30粒/盒', '22', '成都康弘制药有限公司');
INSERT INTO `base_drug` VALUES (146, '速效救心丸', '40mg*50s*3瓶', '盒', '38', '天津中新药业集团有限公司');
INSERT INTO `base_drug` VALUES (147, '羧甲司坦口服溶液', '10ml*0.5g*10支', '盒', '20.89', '北京诚济制药股份有限公司');
INSERT INTO `base_drug` VALUES (148, '缩泉胶囊', '0.3g*60粒', '瓶', '32', '湖南汉森制药');
INSERT INTO `base_drug` VALUES (149, '天王补心丸', '9克*10丸', '盒', '6.7', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (150, '通络祛痛膏', '10片', '盒', '33.18', '河南羚锐制药股份有限公司');
INSERT INTO `base_drug` VALUES (151, '通心络胶囊', '每粒装0.26g', '30粒/盒', '26.38', '石家庄以岭药业股份有限公司');
INSERT INTO `base_drug` VALUES (152, '通宣理肺颗粒', '9g*10袋', '盒', '14.74', '湖北午时');
INSERT INTO `base_drug` VALUES (153, '通宣理肺丸', '6g*10丸', '盒', '5.1', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (154, '尪痹片', '每片重0.5g', '48片/盒', '33.12', '辽宁上药好护士药业（集团）有限公司');
INSERT INTO `base_drug` VALUES (155, '维生素B12注射液', '1ml*10支*0.5mg', '盒', '3.5', '河南润弘制药');
INSERT INTO `base_drug` VALUES (156, '维生素B1注射液', '0.1g*10支', '盒', '0.43', '天津金耀氨基酸有限公司');
INSERT INTO `base_drug` VALUES (157, '维生素B2片', '5mg*100片', '瓶', '5.3', '天津力生制药股份有限公司');
INSERT INTO `base_drug` VALUES (158, '维生素B6注射液', '0.1g*2ml*10支', '盒', '0', '山东方明');
INSERT INTO `base_drug` VALUES (159, '维生素C注射液', '1g*5ml*5支', '盒', '3.395', '成都倍特');
INSERT INTO `base_drug` VALUES (160, '维生素D2软胶囊', '0.125mg*10粒', '盒', '48.49', '大连水产药业有限公司');
INSERT INTO `base_drug` VALUES (161, '稳心颗粒', '5g*9d', '盒', '26.24', '山东步长制药有限公司');
INSERT INTO `base_drug` VALUES (162, '乌鸡白凤丸', '9g*10w', '盒', '21.65', '药都制药有限公司');
INSERT INTO `base_drug` VALUES (163, '乌灵胶囊', '0.33g*9s*3板*2袋', '盒', '58.71', '浙江佐力药业股份有限公司');
INSERT INTO `base_drug` VALUES (164, '腺苷钴胺片', '0.25mg*100片', '盒', '31.78', '石药集团欧意药业');
INSERT INTO `base_drug` VALUES (165, '香砂平胃颗粒', '每袋装5g（无蔗糖）', '6袋/盒', '18.77', '昆明中药厂有限公司');
INSERT INTO `base_drug` VALUES (166, '香砂养胃丸', '9g10袋', '盒', '11.5', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (167, '逍遥丸', '9克*10丸', '盒', '6.05', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (168, '消渴丸', '210丸', '盒', '35', '广州白云山制药股份有限公司广州白云山制药总厂');
INSERT INTO `base_drug` VALUES (169, '消银颗粒', '3.5g', '12袋/盒', '25.26', '陕西康惠制药股份有限公司');
INSERT INTO `base_drug` VALUES (170, '硝苯地平缓释片', '20mg*45片', '瓶', '17.78', '德州博诚制药有限公司');
INSERT INTO `base_drug` VALUES (171, '硝酸咪康唑乳膏', '15g', '支', '4.9', '福元药业');
INSERT INTO `base_drug` VALUES (172, '硝酸咪康唑栓', '0.2g7粒', '盒', '2.2857', NULL);
INSERT INTO `base_drug` VALUES (173, '硝酸异山梨酯片', '5mg*100p', '瓶', '8', '天津太平洋制药有限公司');
INSERT INTO `base_drug` VALUES (174, '小儿宝泰康颗粒', '每袋装2.6克', '18袋/盒', '24.93', '健民集团叶开泰国药（随州）有限公司');
INSERT INTO `base_drug` VALUES (175, '小儿柴桂退热颗粒', '2.5*16袋', '盒', '39.6', '葵花药业');
INSERT INTO `base_drug` VALUES (176, '小儿肺咳颗粒', '2g*9袋', '盒', '27.08', '长春人民药业');
INSERT INTO `base_drug` VALUES (177, '小儿肺热咳喘颗粒', '4g', '16袋/盒', '45.62', '海南葫芦娃药业集团股份有限公司');
INSERT INTO `base_drug` VALUES (178, '小儿热速清颗粒', '6g*10d', '盒', '26.65', '江西倍肯');
INSERT INTO `base_drug` VALUES (179, '小金胶囊', '每粒装0.35克', '9粒/盒', '30.19', '健民药业集团股份有限公司');
INSERT INTO `base_drug` VALUES (180, '心可舒片', '0.31*72片', '盒', '30.66', '山东沃华医药科技有限公司');
INSERT INTO `base_drug` VALUES (181, '醒脾养儿颗粒', '2g*18袋', '盒', '25.5', '贵州建兴');
INSERT INTO `base_drug` VALUES (182, '玄麦甘桔颗粒', '10g*20袋', '盒', '28', '四川彩虹');
INSERT INTO `base_drug` VALUES (183, '血栓心脉宁胶囊', '每粒装0.5g', '40粒/盒', '0', NULL);
INSERT INTO `base_drug` VALUES (184, '血脂康胶囊', '每粒装0.3g', '24粒/盒', '26.08', '北京北大维信生物科技有限公司');
INSERT INTO `base_drug` VALUES (185, '盐酸氨溴索口服液', '30mg*15支', '盒', '9', '国药');
INSERT INTO `base_drug` VALUES (186, '盐酸倍他司汀片', '4mg*10片', '盒', '23.59', '东普恒久远');
INSERT INTO `base_drug` VALUES (187, '盐酸吡格列酮片', '15mg*28片', '盒', '2.1429', '烟台');
INSERT INTO `base_drug` VALUES (188, '盐酸氟桂利嗪胶囊', '5mg*20粒', '盒', '21.16', '西安杨森制药有限公司');
INSERT INTO `base_drug` VALUES (189, '盐酸雷尼替丁胶囊', '0.15g*30粒', '瓶', '9.5', '上海衡山');
INSERT INTO `base_drug` VALUES (190, '盐酸利多卡因注射液', '5ml', '盒', '2.3', NULL);
INSERT INTO `base_drug` VALUES (191, '盐酸莫雷西嗪片', '50mg*20s*2板', '盒', '0', '丹东医创药业有限公司');
INSERT INTO `base_drug` VALUES (192, '盐酸羟甲唑啉喷雾剂', '10ml：5mg2瓶/盒', '盒', '36.8', '深圳大佛药业股份有限公司');
INSERT INTO `base_drug` VALUES (193, '盐酸消旋山莨菪碱注射液', '10mg*10支', '盒', '2.38', '杭州民生药业集团有限公司');
INSERT INTO `base_drug` VALUES (194, '养血清脑颗粒', '4g*15袋', '盒', '31.29', '天津天士力制药股份有限公司');
INSERT INTO `base_drug` VALUES (195, '养阴清肺丸', '9克*10丸', '盒', '24.4', '山西天生制药有限责任公司');
INSERT INTO `base_drug` VALUES (196, '叶酸片', '5mg100片', '瓶', '19.8', '甘肃兰药药业有限公司');
INSERT INTO `base_drug` VALUES (197, '一清颗粒', '7.5g*12袋', '盒', '17', '郑州福瑞堂制药');
INSERT INTO `base_drug` VALUES (198, '益母草颗粒', '15g*10袋', '盒', '17', '上海宝龙安庆药业有限公司');
INSERT INTO `base_drug` VALUES (199, '益气和胃胶囊', '0.5g*36粒', '盒', '44.8', '合肥立法制药有限公司');
INSERT INTO `base_drug` VALUES (200, '益气维血片', '0.57g*24片', '盒', '30.77', '广东红珊瑚');

SET FOREIGN_KEY_CHECKS = 1;
