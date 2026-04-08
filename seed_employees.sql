-- Seed employees from local JSON into Supabase
-- Run this in Supabase SQL Editor

TRUNCATE TABLE employees;

INSERT INTO employees (matricule, nom, prenom, responsable, service, poste, last_seen) VALUES
('MD0268', 'AIT OBIDAT', 'MAROUANE', 'Bennis Adnan', 'export', 'Chargé Exp', '2026-02-01'),
('MD1282', 'AIT-OUSFAI', 'FATIMA', 'Bennis Adnan', 'recouvrement', 'Assistant', NULL),
('MD1826', 'ASSAKKOUR', 'FOUZIA', 'Bennis Adnan', 'export', 'Chargé Export', '2025-12-25'),
('MD0311', 'BELFALAH', 'AMINE', 'Bennis Adnan', 'commercial', 'Assist-Com', '2025-12-25'),
('MD1842', 'DAHANE', 'ALAE-EDDINE', 'Bennis Adnan', 'commercial', 'Big data', NULL),
('MB0631', 'EL MAMOURI', 'AYOUB', 'Bennis Adnan', 'informatique', 'Informatiq', '2025-12-25'),
('MD1451', 'EL OUARDI', 'BILAL', 'Bennis Adnan', 'commercial', 'Chargé Com', '2025-12-25'),
('MB5206', 'EL YOUSSRI', 'MOHAMMED CHARAF', 'Bennis Adnan', 'commercial', 'Repres-Com', NULL),
('MD1040', 'ELASLI', 'BTISSAM', 'Bennis Adnan', 'import', 'Chargé Log', '2025-12-25'),
('MD1008', 'EZ-ZYN', 'HAFSA', 'Bennis Adnan', 'commercial', 'Chargé de', NULL),
('MI0537', 'LAMRABTI', 'YASSINE', 'Bennis Adnan', 'export', 'Chargé Com', '2025-12-25'),
('MC4444', 'SAOUD', 'NASRALLAH', 'Bennis Adnan', 'commercial', 'Repres-Com', NULL),
('MD0371', 'TOUNCHIBIN', 'LEILA', 'Bennis Adnan', 'recouvrement', 'GMS', NULL),
('MD1279', 'ASBAYOU', 'ASMAE', 'Bennis khalid', 'Tresorerie', 'Chargé Trésorie', '2025-12-25'),
('MB0031', 'BADROUN', 'RACHID', 'Bennis khalid', 'ressources humains', 'Chargé RH', '2025-12-25'),
('MD1828', 'BELQADI', 'SAFA', 'Bennis khalid', 'informatique', 'Big data', '2025-12-25'),
('MD1730', 'BOULAFDAM', 'FATIMA-ZAHRA', 'Bennis khalid', 'comptabilite', 'Comptable', NULL),
('MD1871', 'EL JAKOUNE', 'ABDOULHAKIM', 'Bennis khalid', 'comptabilite', 'Comptable', NULL),
('MD0282', 'IHSSANE', 'OUALID', 'Bennis khalid', 'informatique', 'Chargé Ach', '2025-12-25'),
('MD1644', 'MADANI', 'MERYEM', 'Bennis khalid', 'Methode', 'Methode', '2025-12-25'),
('MD0742', 'MESKINI', 'NOURA', 'Bennis khalid', 'comptabilite', 'Attaché Di', '2025-12-25'),
('MD1571', 'SEDDATI', 'MOHAMMED', 'Bennis khalid', 'comptabilite', 'Comptable', '2025-12-25'),
('MC0734', 'SEMLALI', 'HAMZA', 'Bennis khalid', 'Methode', 'Opérateur', NULL),
('MD0745', 'TAHOUM', 'SOLTANA', 'Bennis khalid', 'ressources humains', 'Chargé RH', '2025-12-25'),
('MB0316', 'TAIBA', 'ZOUHAIR', 'Bennis khalid', 'informatique', 'Attaché Di', NULL),
('MD1632', 'ALLAM', 'SAAD', 'Bennis Younes', 'Infographie', 'Chargé Infographie', '2025-12-25'),
('MD1376', 'EL AIDOUNI', 'HIND', 'Bennis Younes', 'recouvrement', 'Chargé Inf', NULL),
('MB4663', 'EZ-ZAFRANI', 'REDOUANE', 'Bennis Younes', 'achat', 'Chargé Sto', NULL),
('MB1220', 'HARMOUCH', 'FADOUA', 'Bennis Younes', 'recouvrement', 'Force de V', '2025-12-25'),
('MR0002', 'HMAMOU', 'HAMID', 'Bennis Younes', 'Infographie', 'Chargé Inf', '2025-12-25'),
('MD1386', 'ZOUHRI', 'ZAINAB', 'Bennis Younes', 'Methode', 'Assist ADV', NULL);
