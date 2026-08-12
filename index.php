<?php
// ==============================================================================
// FANTASTRATEGY EXCEL SUPREME EDITION ⚡ - AUCTION LIVE ENGINE 2026/2027
// ==============================================================================

// Gestione percorso dinamico: rileva se siamo su Windows (XAMPP) o Linux (Render)
if (strtoupper(substr(PHP_OS, 0, 3)) === 'WIN') {
    $tmpDir = sys_get_temp_dir();
    $dataFile = $tmpDir . DIRECTORY_SEPARATOR . 'fantacalcio_supreme_db.json';
} else {
    $dataFile = '/tmp/fantacalcio_supreme_db.json';
}

$defaultFile = __DIR__ . '/fantacalcio_supreme_db.json';

// Se il file temporaneo non esiste ancora, prova a copiarlo o a crearlo
if (!file_exists($dataFile)) {
    if (file_exists($defaultFile)) {
        copy($defaultFile, $dataFile);
    } else {
        $initialData = [
            'lega1' => [
                'nome' => 'Lega 1', 
                'budget_iniziale' => 500, 
                'budget_ruoli' => ['P' => 40, 'D' => 80, 'C' => 150, 'A' => 230],
                'giocatori' => [], 
                'coppie' => []
            ]
        ];
        file_put_contents($dataFile, json_encode($initialData, JSON_PRETTY_PRINT));
    }
}

// Lettura e verifica dati
$fileContent = file_exists($dataFile) ? file_get_contents($dataFile) : '';
$data = json_decode($fileContent, true);

if (!is_array($data)) {
    $data = [
        'lega1' => [
            'nome' => 'Lega 1', 
            'budget_iniziale' => 500, 
            'budget_ruoli' => ['P' => 40, 'D' => 80, 'C' => 150, 'A' => 230],
            'giocatori' => [], 
            'coppie' => []
        ]
    ];
}

$currentLega = $_GET['lega'] ?? 'lega1';
if (!array_key_exists($currentLega, $data)) { $currentLega = 'lega1'; }

// EXPORT BACKUP IN FORMATO JSON / CSV
if (isset($_GET['action']) && $_GET['action'] === 'export_backup') {
    $type = $_GET['type'] ?? 'json';
    if ($type === 'csv') {
        header('Content-Type: text/csv; charset=utf-8');
        header('Content-Disposition: attachment; filename=fantastrategy_export_' . date('Ymd_His') . '.csv');
        $output = fopen('php://output', 'w');
        fputcsv($output, ['Nome', 'Ruolo', 'Squadra', 'Fascia', 'Stato', 'Prezzo Acquisto', 'Budget Max', 'FMV', 'Note']);
        foreach ($data[$currentLega]['giocatori'] as $p) {
            fputcsv($output, [
                $p['nome'] ?? '',
                $p['ruolo'] ?? '',
                $p['squadra'] ?? '',
                $p['fascia'] ?? '',
                $p['stato'] ?? 'da_comprate',
                $p['prezzo_acquisto'] ?? 0,
                $p['budget_max'] ?? 1,
                $p['fmv'] ?? 0,
                $p['note'] ?? ''
            ]);
        }
        fclose($output);
        exit;
    } else {
        header('Content-Type: application/json; charset=utf-8');
        header('Content-Disposition: attachment; filename=fantastrategy_backup_' . date('Ymd_His') . '.json');
        echo json_encode($data, JSON_PRETTY_PRINT);
        exit;
    }
}

// Helper per pulire le intestazioni
function cleanHeaderStr($str) {
    $str = mb_strtolower(trim((string)$str), 'UTF-8');
    $str = str_replace(['à', 'è', 'é', 'ì', 'ò', 'ù', '.', ' '], ['a', 'e', 'e', 'i', 'o', 'u', '', ''], $str);
    return $str;
}

// Convertitore di coordinate colonna Excel
function excelColumnToIndex($colRef) {
    $colRef = strtoupper(preg_replace('/[0-9]/', '', $colRef));
    $length = strlen($colRef);
    $index = 0;
    for ($i = 0; $i < $length; $i++) {
        $index = $index * 26 + (ord($colRef[$i]) - ord('A') + 1);
    }
    return $index - 1;
}

// ==============================================================================
// DATABASE TATTICO COMPLETO DELLE 20 SQUADRE
// ==============================================================================
$tatticheSquadre = [
    'ATA' => [
        'allenatore' => 'M. Sarri',
        'rigoristi' => ['Scamacca', 'Samardzic', 'Ederson D.S.'],
        'punizioni' => ['Samardzic', 'Gaetano'],
        'angoli' => ['Samardzic', 'Gaetano', 'Bernasconi'],
        'ballottaggi' => ['Scamacca (51%) vs Krstovic (49%)', 'Gaetano (51%) vs De Roon (49%)', 'Zappacosta (55%) vs Bellanova (45%)', 'Kolasinac (55%) vs Hien (45%)', 'Samardzic (51%) vs Pasalic (49%)', 'Bernasconi (51%) vs Ahanor (49%)'],
        'valorizzati' => ['Scalvini' => 'Difesa di Sarri solida', 'Raspadori' => 'Adattabile anche al centro, esterni prolifici'],
        'penalizzati' => ['De Roon' => 'Posizionamento più arretrato', 'Kossounou' => 'Parte indietro nelle gerarchie'],
        'nascosti' => ['Samardzic'],
        'giovani' => ['Bernasconi']
    ],
    'BOL' => [
        'allenatore' => 'D. Tedesco',
        'rigoristi' => ['Orsolini', 'Dovbyk', 'Bernardeschi'],
        'punizioni' => ['Orsolini', 'Bernardeschi'],
        'angoli' => ['Orsolini', 'Bernardeschi', 'Miranda J.'],
        'ballottaggi' => ['Vitik (51%) vs Heggem (49%)', 'Zortea (51%) vs Holm (49%)', 'Bernardeschi (60%) vs Odgaard (40%)', 'Pobega (55%) vs Moro N. (45%)'],
        'valorizzati' => ['Rowe' => 'Caratteristiche in linea con gli esterni', 'Orsolini' => 'Perno offensivo e perno bonus'],
        'penalizzati' => ['Dallinga' => 'Indietro nelle gerarchie', 'Lucumì' => 'Soffre sistema di gioco elaborato'],
        'nascosti' => ['Bernardeschi'],
        'giovani' => ['Dominguez B.']
    ],
    'CAG' => [
        'allenatore' => 'F. Pisacane',
        'rigoristi' => ['Fazzini', 'Mina', 'Borrelli'],
        'punizioni' => ['Fazzini', 'Maldini'],
        'angoli' => ['Fazzini', 'Obert', 'Maldini'],
        'ballottaggi' => ['Mina (51%) vs Deiola (49%)', 'Zè Pedro (55%) vs Zappa (45%)', 'Borrelli (55%) vs Mendy P. (45%)'],
        'valorizzati' => ['Obert' => 'Buoni cross dai piazzati', 'Fazzini' => 'Voluto dal mister, propensione al bonus'],
        'penalizzati' => ['Zappa' => 'Sistema troppo difensivo'],
        'nascosti' => ['Romano'],
        'giovani' => ['Mendy P.']
    ],
    'COM' => [
        'allenatore' => 'C. Fabregas',
        'rigoristi' => ['Da Cunha', 'Douvikas', 'Baturina'],
        'punizioni' => ['Paz N.', 'Baturina'],
        'angoli' => ['Paz N.', 'Baturina', 'Da Cunha'],
        'ballottaggi' => ['Perrone (60%) vs Milla (40%)', 'Diao (60%) vs Rodriguez Je. (40%)', 'Couto (51%) vs Smolcic I. (49%)', 'Valle (51%) vs Kaiki (49%)', 'Ramon (55%) vs Kempf (45%)'],
        'valorizzati' => ['Da Cunha' => 'Duttile tatticamente', 'Douvikas' => 'Manovra offensiva crea spazi'],
        'penalizzati' => ['Morata' => 'Non ancora integrato', 'Valle' => 'Molto turnover nel suo ruolo'],
        'nascosti' => ['Diao'],
        'giovani' => ['Liberali']
    ],
    'FIO' => [
        'allenatore' => 'F. Grosso',
        'rigoristi' => ['Kean', 'Gudmundsson A.', 'Mastantuono'],
        'punizioni' => ['Mastantuono', 'Fagioli'],
        'angoli' => ['Fagioli', 'Mastantuono', 'Mandragora'],
        'ballottaggi' => ['Jimenez A. (51%) vs Joao Mario (49%)', 'Viery (51%) vs Ranieri L. (49%)', 'Ndour (51%) vs Gudmundsson A. (49%)'],
        'valorizzati' => ['Jimenez A.' => 'Usato sia come terzino che ala', 'Atta' => 'Mezz\'ali di Grosso molto prolifiche'],
        'penalizzati' => ['Brescianini' => 'Indietro nelle gerarchie', 'Ranieri L.' => 'Titolarità non garantita'],
        'nascosti' => ['Ndour'],
        'giovani' => ['Oulai']
    ],
    'FRO' => [
        'allenatore' => 'M. Alvini',
        'rigoristi' => ['Calò', 'Ghedjemis', 'Kvernadze'],
        'punizioni' => ['Calò', 'Ghedjemis'],
        'angoli' => ['Calò', 'Kvernadze', 'Ghedjemis'],
        'ballottaggi' => ['El Azzouzi (51%) vs Gelli (49%)', 'Palmisani (51%) vs Desplanches (49%)', 'Zerbin (51%) vs Kvernadze (49%)', 'Koutsoupias (51%) vs Hasa (49%)'],
        'valorizzati' => ['Ghedjemis' => 'Uomo chiave nel gioco', 'Calò' => 'Calcia tutti i piazzati e rigorista'],
        'penalizzati' => ['Monterisi' => 'Propensione al malus'],
        'nascosti' => ['Zerbin'],
        'giovani' => ['Fini']
    ],
    'GEN' => [
        'allenatore' => 'D. De Rossi',
        'rigoristi' => ['Colombo', 'Ostigard', 'Vitinha O.'],
        'punizioni' => ['Baldanzi', 'Messias'],
        'angoli' => ['Baldanzi', 'Messias', 'Frendrup'],
        'ballottaggi' => ['Vitinha O. (51%) vs Meichtry (49%)', 'Baldanzi (51%) vs Traorè Hj. (49%)', 'Mitaj (51%) vs Ellertsson (49%)'],
        'valorizzati' => ['Ostigard' => 'Prolifico sui piazzati', 'Baldanzi' => 'Fantasista e fulcro di gioco'],
        'penalizzati' => ['Martin' => 'Titolarità non garantita'],
        'nascosti' => ['Marcandalli'],
        'giovani' => ['Puczka']
    ],
    'INT' => [
        'allenatore' => 'C. Chivu',
        'rigoristi' => ['Calhanoglu', 'Zielinski', 'Martinez L.'],
        'punizioni' => ['Dimarco', 'Calhanoglu'],
        'angoli' => ['Calhanoglu', 'Dimarco', 'Zielinski'],
        'ballottaggi' => ['Zielinski (55%) vs Sucic P. (45%)', 'Diouf (51%) vs Luis Henrique (49%)', 'Martinez Jo. (55%) vs Provedel (45%)', 'Stones (55%) vs Bisseck (45%)'],
        'valorizzati' => ['Martinez L.' => 'Fulcro del gioco', 'Dimarco' => 'Ruolo molto offensivo con Chivu'],
        'penalizzati' => ['Frattesi' => 'Non integrato nel gioco', 'Luis Henrique' => 'Indietro nelle gerarchie'],
        'nascosti' => ['Diouf'],
        'giovani' => ['Esposito F.P.']
    ],
    'JUV' => [
        'allenatore' => 'L. Spalletti',
        'rigoristi' => ['Kolo Muani', 'Locatelli', 'Yildiz'],
        'punizioni' => ['Yildiz', 'Cambiaso'],
        'angoli' => ['Yildiz', 'Cambiaso', 'Locatelli'],
        'ballottaggi' => ['Alajbegovic (51%) vs Conceicao (49%)'],
        'valorizzati' => ['Yildiz' => 'Riferimento principale offensivo', 'McKennie' => 'Prolifico negli inserimenti'],
        'penalizzati' => ['Koopmeiners' => 'Posizionamento molto arretrato', 'Zhegrova' => 'Minutaggio ridotto'],
        'nascosti' => ['Kelly L.'],
        'giovani' => ['Ekhator']
    ],
    'LAZ' => [
        'allenatore' => 'G. Gattuso',
        'rigoristi' => ['Zaccagni', 'Taylor K.', 'Cataldi'],
        'punizioni' => ['Zaccagni', 'Cataldi'],
        'angoli' => ['Zaccagni', 'Taylor K.', 'Rovella'],
        'ballottaggi' => ['Rovella (55%) vs Cataldi (45%)', 'Isaksen (60%) vs Cancellieri (40%)', 'Pedraza (51%) vs Tavares N. (49%)', 'Romagnoli (60%) vs Provstgaard (40%)'],
        'valorizzati' => ['Taylor K.' => 'Inserimenti richiesti dal mister', 'Zaccagni' => 'Fantasista e fulcro del gioco'],
        'penalizzati' => ['Rovella' => 'Propensione ai cartellini nel gioco di Gattuso'],
        'nascosti' => ['Pedraza'],
        'giovani' => ['Provstgaard']
    ],
    'LEC' => [
        'allenatore' => 'E. Di Francesco',
        'rigoristi' => ['Geubbels', 'Stulic', 'Pierotti'],
        'punizioni' => ['Gallo', 'Pierotti'],
        'angoli' => ['Gallo', 'Pierotti', 'Berisha M.'],
        'ballottaggi' => ['Berisha M. (51%) vs Ngom (49%)', 'Geubbels (51%) vs Stulic (49%)'],
        'valorizzati' => ['N\'Dri' => 'Ali fondamentali per Di Francesco', 'Coulibaly L.' => 'Fulcro del centrocampo'],
        'penalizzati' => ['Fofana Sa.' => 'Titolarità non garantita'],
        'nascosti' => ['Gandelman'],
        'giovani' => ['Tiago Gabriel']
    ],
    'MIL' => [
        'allenatore' => 'R. Amorim',
        'rigoristi' => ['Ramos G.', 'Nkunku', 'Pulisic'],
        'punizioni' => ['Modric', 'Pulisic'],
        'angoli' => ['Modric', 'Pulisic', 'Bartesaghi'],
        'ballottaggi' => ['Leao (55%) vs Nkunku (45%)'],
        'valorizzati' => ['Pulisic' => 'Fantasista adattabile ovunque', 'Ramos G.' => 'Fulcro del gioco di Amorim'],
        'penalizzati' => ['Tomori' => 'Non adatto all\'impostazione richiesta', 'Ricci S.' => 'Troppo campo da coprire'],
        'nascosti' => ['Nkunku'],
        'giovani' => ['Bartesaghi']
    ],
    'MON' => [
        'allenatore' => 'I. Jurić',
        'rigoristi' => ['Pessina', 'Cutrone', 'Petagna'],
        'punizioni' => ['Colpani', 'Pessina'],
        'angoli' => ['Pessina', 'Colpani', 'Ciurria'],
        'ballottaggi' => ['Cutrone (51%) vs Petagna (49%)', 'Colpani (55%) vs Ciurria (45%)', 'Mota (55%) vs Ciurria (45%)', 'Akinsanmiro (51%) vs Colombo L. (49%)'],
        'valorizzati' => ['Mota' => 'Leader d\'attacco', 'Birindelli' => 'Spinta costante sulle fasce'],
        'penalizzati' => ['Colpani' => 'Troppo discontinuo'],
        'nascosti' => ['Thiam'],
        'giovani' => ['Bakoune']
    ],
    'NAP' => [
        'allenatore' => 'M. Allegri',
        'rigoristi' => ['De Bruyne', 'Hojlund', 'Lukaku'],
        'punizioni' => ['De Bruyne', 'Politano'],
        'angoli' => ['De Bruyne', 'Politano', 'Neres'],
        'ballottaggi' => ['Santos A. (55%) vs Politano (45%)', 'Marin R. (51%) vs Olivera (49%)'],
        'valorizzati' => ['Santos A.' => 'Prolifico in contropiede', 'McTominay' => 'Prolifico sui piazzati e inserimenti'],
        'penalizzati' => ['Gilmour' => 'Indietro nelle gerarchie'],
        'nascosti' => ['Olivera'],
        'giovani' => ['Vergara']
    ],
    'PAR' => [
        'allenatore' => 'C. Cuesta',
        'rigoristi' => ['Pellegrino M.', 'Valeri', 'Bernabè'],
        'punizioni' => ['Bernabè', 'Nicolussi Caviglia'],
        'angoli' => ['Bernabè', 'Nicolussi Caviglia', 'Valeri'],
        'ballottaggi' => ['Nicolussi Caviglia (51%) vs Sorensen O. (49%)', 'Tourè E. (55%) vs Diallo O. (45%)', 'Corvi (55%) vs Daffara (45%)', 'Frigan (51%) vs Ondrejka (49%)'],
        'valorizzati' => ['Bernabè' => 'Riferimento tecnico del centrocampo', 'Pellegrino M.' => 'Terminale sulle palle alte'],
        'penalizzati' => ['Sorensen O.' => 'Indietro nelle gerarchie'],
        'nascosti' => ['Delprato'],
        'giovani' => ['Elphege']
    ],
    'ROM' => [
        'allenatore' => 'G. Gasperini',
        'rigoristi' => ['Malen', 'Dybala', 'Soulè'],
        'punizioni' => ['Dybala', 'Soulè'],
        'angoli' => ['Dybala', 'Soulè', 'Wesley'],
        'ballottaggi' => ['Koulierakis (55%) vs Hermoso (45%)', 'Soulè (51%) vs Pellegrini Lo. (49%)', 'Dybala (55%) vs Castro S. (45%)'],
        'valorizzati' => ['Malen' => 'Pressing gli porta palloni puliti', 'Wesley' => 'Esterno a tutto campo valorizzato'],
        'penalizzati' => ['Vaz' => 'Troppa concorrenza'],
        'nascosti' => ['Cristante'],
        'giovani' => ['Pisilli']
    ],
    'SAS' => [
        'allenatore' => 'A. Aquilani',
        'rigoristi' => ['Berardi', 'Pinamonti', 'Laurientè'],
        'punizioni' => ['Berardi', 'Laurientè'],
        'angoli' => ['Berardi', 'Laurientè', 'Doig'],
        'ballottaggi' => ['Walukiewicz (51%) vs Candè (49%)', 'Muric (51%) vs Turati (49%)', 'Lipani (51%) vs Konè I. (49%)', 'Pinamonti (60%) vs Bowie (40%)'],
        'valorizzati' => ['Laurientè' => 'Attacco della profondità', 'Berardi' => 'Riferimento e rigorista top'],
        'penalizzati' => ['Muric' => 'Soffre la costruzione dal basso', 'Walukiewicz' => 'Soffre costruzione dal basso'],
        'nascosti' => ['Matic'],
        'giovani' => ['Lipani']
    ],
    'TOR' => [
        'allenatore' => 'I. Abate',
        'rigoristi' => ['Vlasic', 'Kulenovic', 'Zapata D.'],
        'punizioni' => ['Vlasic', 'Oristanio'],
        'angoli' => ['Vlasic', 'Oristanio', 'Ilic'],
        'ballottaggi' => ['Oristanio (51%) vs Adams C. (49%)', 'Comuzzo (51%) vs Comert (49%)', 'Fitz-Jim (51%) vs Gineitis (49%)'],
        'valorizzati' => ['Simeone' => 'Utile nel gioco aereo per gli esterni', 'Vlasic' => 'Svaria tra le lines'],
        'penalizzati' => ['Biraghi' => 'Soffre la difesa alta'],
        'nascosti' => ['Oristanio'],
        'giovani' => ['Njie']
    ],
    'UDI' => [
        'allenatore' => 'K. Runjaic',
        'rigoristi' => ['Davis K.', 'Solet', 'Ekkelenkamp'],
        'punizioni' => ['Zaniolo', 'Ekkelenkamp'],
        'angoli' => ['Zaniolo', 'Ekkelenkamp', 'Vojvoda'],
        'ballottaggi' => ['Piotrowski (55%) vs Unai Gomez (45%)'],
        'valorizzati' => ['Davis K.' => 'Presenza fisica per le ripartenze', 'Ekkelenkamp' => 'Riferimento e inserimenti'],
        'penalizzati' => ['Karlstrom' => 'Posizionamento molto arretrato'],
        'nascosti' => ['Kristensen T.'],
        'giovani' => ['Miller L.']
    ],
    'VEN' => [
        'allenatore' => 'G. Stroppa',
        'rigoristi' => ['Adams A.', 'Rrahmani Al.', 'Yeboah J.'],
        'punizioni' => ['Busio', 'Basic'],
        'angoli' => ['Busio', 'Perez K.', 'Basic'],
        'ballottaggi' => ['Correia T. (51%) vs Hainaut (49%)', 'Adams A. (55%) vs Rrahmani Al. (45%)'],
        'valorizzati' => ['Yeboah J.' => 'Dribbling e profondità', 'Busio' => 'Uomo chiave nel gioco del mister'],
        'penalizzati' => ['Franjic' => 'Troppa concorrenza'],
        'nascosti' => ['Perez K.'],
        'giovani' => ['Dagasso']
    ]
];

// ==============================================================================
// CLASSIFICAZIONE DIFFICOLTÀ SQUADRE
// ==============================================================================
$teamCategories = [
    'NAP' => 3, 'INT' => 3, 'MIL' => 3, 'ROM' => 3, 'COM' => 3, 'JUV' => 3,
    'LAZ' => 2, 'ATA' => 2, 'FIO' => 2, 'BOL' => 2,
    'SAS' => 1, 'TOR' => 1, 'EMP' => 1, 'CAG' => 1, 'MON' => 1, 
    'PAR' => 1, 'UDI' => 1, 'VEN' => 1, 'FRO' => 1, 'GEN' => 1, 'LEC' => 1
];

$teamFixtures = [
    'ROM' => ['FIO', 'LEC', 'ATA', 'TOR', 'INT', 'COM', 'GEN', 'NAP', 'CAG', 'UDI', 'SAS', 'PAR', 'MON', 'BOL', 'LAZ', 'JUV', 'EMP', 'VEN', 'MIL', 'ATA', 'UDI', 'MON', 'TOR', 'PAR', 'SAS', 'VEN', 'JUV', 'GEN', 'LEC', 'BOL', 'INT', 'LAZ', 'EMP', 'NAP', 'FIO', 'MIL', 'COM', 'CAG'],
    'COM' => ['UDI', 'NAP', 'GEN', 'PAR', 'FIO', 'ROM', 'LAZ', 'MON', 'VER', 'INT', 'CAG', 'JUV', 'LEC', 'BOL', 'MIL', 'VEN', 'EMP', 'ATA', 'TOR', 'NAP', 'UDI', 'MON', 'VER', 'BOL', 'MIL', 'PAR', 'GEN', 'JUV', 'FIO', 'CAG', 'LEC', 'LAZ', 'TOR', 'INT', 'ATA', 'EMP', 'ROM', 'VEN'],
    'INT' => ['TOR', 'LEC', 'ATA', 'MON', 'ROM', 'CAG', 'JUV', 'NAP', 'VEN', 'COM', 'BOL', 'LAZ', 'FIO', 'MIL', 'PAR', 'SAS', 'EMP', 'GEN', 'UDI', 'ATA', 'MON', 'LEC', 'SAS', 'PAR', 'JUV', 'CAG', 'MIL', 'NAP', 'FIO', 'BOL', 'ROM', 'LAZ', 'TOR', 'COM', 'GEN', 'VEN', 'EMP', 'UDI'],
    'JUV' => ['PAR', 'VEN', 'ROM', 'EMP', 'MIL', 'CAG', 'INT', 'LAZ', 'BOL', 'TOR', 'FIO', 'COM', 'NAP', 'GEN', 'UDI', 'ROM', 'MON', 'ATA', 'SAS', 'VEN', 'PAR', 'EMP', 'CAG', 'MIL', 'INT', 'LAZ', 'ROM', 'COM', 'TOR', 'BOL', 'FIO', 'GEN', 'NAP', 'UDI', 'SAS', 'MON', 'ATA', 'VEN'],
    'MIL' => ['TOR', 'PAR', 'LAZ', 'VEN', 'JUV', 'LEC', 'FIO', 'UDI', 'BOL', 'NAP', 'MON', 'CAG', 'EMP', 'INT', 'COM', 'GEN', 'ROM', 'ATA', 'SAS', 'TOR', 'PAR', 'LAZ', 'VEN', 'LEC', 'COM', 'FIO', 'INT', 'UDI', 'BOL', 'NAP', 'MON', 'CAG', 'EMP', 'SAS', 'GEN', 'ROM', 'ATA', 'JUV'],
    'NAP' => ['VER', 'COM', 'PAR', 'JUV', 'MON', 'COM', 'EMP', 'ROM', 'MIL', 'ATA', 'INT', 'LAZ', 'JUV', 'TOR', 'UDI', 'GEN', 'VEN', 'FIO', 'BOL', 'COM', 'PAR', 'JUV', 'MON', 'EMP', 'MIL', 'ATA', 'INT', 'ROM', 'LAZ', 'TOR', 'UDI', 'GEN', 'JUV', 'ROM', 'VEN', 'FIO', 'BOL', 'VER'],
    'LAZ' => ['VEN', 'UDI', 'MIL', 'VER', 'FIO', 'TOR', 'EMP', 'JUV', 'GEN', 'CAG', 'MON', 'NAP', 'BOL', 'PAR', 'ROM', 'INT', 'ATA', 'LEC', 'COM', 'VEN', 'UDI', 'MIL', 'VER', 'FIO', 'TOR', 'EMP', 'JUV', 'GEN', 'CAG', 'MON', 'NAP', 'ROM', 'BOL', 'PAR', 'INT', 'ATA', 'LEC', 'COM'],
    'ATA' => ['LEC', 'TOR', 'INT', 'FIO', 'COM', 'BOL', 'GEN', 'VEN', 'VER', 'NAP', 'UDI', 'PAR', 'ROM', 'MIL', 'CAG', 'EMP', 'LAZ', 'JUV', 'MON', 'LEC', 'TOR', 'INT', 'FIO', 'COM', 'BOL', 'GEN', 'VEN', 'VER', 'NAP', 'UDI', 'PAR', 'MIL', 'CAG', 'EMP', 'LAZ', 'JUV', 'MON', 'ROM'],
    'FIO' => ['ROM', 'VEN', 'MON', 'ATA', 'LAZ', 'EMP', 'MIL', 'TOR', 'GEN', 'VER', 'JUV', 'LEC', 'INT', 'CAG', 'BOL', 'UDI', 'PAR', 'NAP', 'COM', 'ROM', 'VEN', 'MON', 'ATA', 'LAZ', 'EMP', 'MIL', 'TOR', 'GEN', 'VER', 'JUV', 'LEC', 'INT', 'CAG', 'BOL', 'UDI', 'PAR', 'NAP', 'COM'],
    'BOL' => ['UDI', 'NAP', 'EMP', 'MON', 'ATA', 'PAR', 'LAZ', 'MIL', 'ROM', 'JUV', 'INT', 'VEN', 'FIO', 'COM', 'TOR', 'GEN', 'VER', 'CAG', 'LEC', 'UDI', 'NAP', 'EMP', 'MON', 'COM', 'ATA', 'PAR', 'LAZ', 'MIL', 'ROM', 'JUV', 'INT', 'VEN', 'FIO', 'TOR', 'GEN', 'VER', 'CAG', 'LEC'],
    'TOR' => ['MIL', 'ATA', 'VEN', 'ROM', 'VER', 'LAZ', 'INT', 'FIO', 'CAG', 'JUV', 'MON', 'NAP', 'GEN', 'EMP', 'BOL', 'UDI', 'PAR', 'COM', 'LEC', 'MIL', 'ATA', 'VEN', 'ROM', 'VER', 'LAZ', 'INT', 'FIO', 'CAG', 'JUV', 'MON', 'NAP', 'GEN', 'EMP', 'BOL', 'UDI', 'PAR', 'COM', 'LEC'],
    'GEN' => ['INT', 'VER', 'ROM', 'VEN', 'JUV', 'ATA', 'FIO', 'LAZ', 'PAR', 'COM', 'MON', 'NAP', 'TOR', 'UDI', 'MIL', 'BOL', 'EMP', 'CAG', 'LEC', 'INT', 'VER', 'ROM', 'VEN', 'JUV', 'ATA', 'FIO', 'LAZ', 'PAR', 'COM', 'MON', 'NAP', 'TOR', 'UDI', 'MIL', 'BOL', 'EMP', 'CAG', 'LEC'],
    'SAS' => ['MON', 'PAR', 'JUV', 'NAP', 'ROM', 'MIL', 'INT', 'LAZ', 'FIO', 'BOL', 'TOR', 'GEN', 'UDI', 'EMP', 'CAG', 'VEN', 'LEC', 'COM', 'ATA', 'MON', 'PAR', 'JUV', 'NAP', 'INT', 'ROM', 'MIL', 'LAZ', 'FIO', 'BOL', 'TOR', 'GEN', 'UDI', 'EMP', 'MIL', 'CAG', 'VEN', 'LEC', 'COM'],
    'PAR' => ['JUV', 'MIL', 'NAP', 'COM', 'UDI', 'BOL', 'LEC', 'FIO', 'GEN', 'EMP', 'ATA', 'ROM', 'LAZ', 'VER', 'INT', 'MON', 'VEN', 'TOR', 'CAG', 'JUV', 'MIL', 'NAP', 'COM', 'UDI', 'BOL', 'LEC', 'FIO', 'GEN', 'EMP', 'ATA', 'ROM', 'LAZ', 'VER', 'INT', 'MON', 'VEN', 'TOR', 'CAG'],
    'UDI' => ['BOL', 'LAZ', 'COM', 'PAR', 'INT', 'ROM', 'MIL', 'JUV', 'ATA', 'FIO', 'GEN', 'NAP', 'TOR', 'MON', 'VEN', 'LEC', 'EMP', 'CAG', 'VER', 'BOL', 'LAZ', 'COM', 'PAR', 'INT', 'ROM', 'MIL', 'JUV', 'ATA', 'FIO', 'GEN', 'NAP', 'TOR', 'MON', 'VEN', 'LEC', 'EMP', 'CAG', 'VER'],
    'MON' => ['SAS', 'FIO', 'BOL', 'NAP', 'TOR', 'COM', 'ROM', 'LAZ', 'INT', 'GEN', 'MIL', 'PAR', 'UDI', 'JUV', 'VEN', 'LEC', 'EMP', 'CAG', 'ATA', 'SAS', 'FIO', 'BOL', 'NAP', 'TOR', 'COM', 'ROM', 'LAZ', 'INT', 'GEN', 'MIL', 'PAR', 'UDI', 'JUV', 'VEN', 'LEC', 'EMP', 'CAG', 'ATA'],
    'LEC' => ['ATA', 'INT', 'TOR', 'MIL', 'PAR', 'LAZ', 'FIO', 'ROM', 'GEN', 'COM', 'VEN', 'UDI', 'JUV', 'MON', 'EMP', 'CAG', 'VER', 'NAP', 'SAS', 'ATA', 'INT', 'TOR', 'MIL', 'PAR', 'LAZ', 'FIO', 'ROM', 'GEN', 'COM', 'VEN', 'UDI', 'JUV', 'MON', 'EMP', 'CAG', 'VER', 'NAP', 'SAS'],
    'CAG' => ['ROM', 'COM', 'JUV', 'INT', 'MIL', 'LAZ', 'TOR', 'GEN', 'MON', 'PAR', 'UDI', 'NAP', 'FIO', 'SAS', 'EMP', 'LEC', 'VEN', 'VER', 'ATA', 'ROM', 'COM', 'JUV', 'INT', 'MIL', 'LAZ', 'TOR', 'GEN', 'MON', 'PAR', 'UDI', 'NAP', 'FIO', 'SAS', 'EMP', 'LEC', 'VEN', 'VER', 'ATA'],
    'VEN' => ['LAZ', 'JUV', 'TOR', 'MIL', 'GEN', 'ATA', 'INT', 'LEC', 'COM', 'PAR', 'UDI', 'MON', 'EMP', 'CAG', 'SAS', 'ROM', 'FIO', 'NAP', 'VER', 'LAZ', 'JUV', 'TOR', 'MIL', 'GEN', 'ATA', 'INT', 'LEC', 'COM', 'PAR', 'UDI', 'MON', 'EMP', 'CAG', 'SAS', 'ROM', 'FIO', 'NAP', 'VER'],
    'EMP' => ['MON', 'ROM', 'JUV', 'LAZ', 'FIO', 'NAP', 'PAR', 'UDI', 'MON', 'GEN', 'COM', 'VEN', 'SAS', 'CAG', 'LEC', 'TOR', 'MIL', 'INT', 'ATA', 'MON', 'ROM', 'JUV', 'LAZ', 'FIO', 'NAP', 'PAR', 'UDI', 'MON', 'GEN', 'COM', 'VEN', 'SAS', 'CAG', 'LEC', 'TOR', 'MIL', 'INT', 'ATA']
];

$teamCalendars = [];
foreach ($teamFixtures as $teamCode => $fixtures) {
    $teamCalendars[$teamCode] = array_map(fn($opp) => $teamCategories[$opp] ?? 1, $fixtures);
}

// Helper per estrarre statistiche squadra
function getTeamCalendarStats($teamCode) {
    global $teamCalendars;
    $cal = $teamCalendars[$teamCode] ?? array_fill(0, 38, 2);
    $counts = array_count_values($cal);
    return [
        'facili' => $counts[1] ?? 18,
        'medie' => $counts[2] ?? 12,
        'difficili' => $counts[3] ?? 8
    ];
}

// PARSER XLSX NATIVO
function parseXlsxNative($filePath) {
    if (!class_exists('ZipArchive')) {
        return ['error' => true, 'message' => 'L\'estensione PHP "zip" non è abilitata!'];
    }

    $zip = new ZipArchive();
    if ($zip->open($filePath) !== true) {
        return ['error' => true, 'message' => 'Impossibile aprire il file .xlsx.'];
    }

    $sharedStrings = [];
    if (($ssXml = $zip->getFromName('xl/sharedStrings.xml')) !== false) {
        $xml = simplexml_load_string($ssXml);
        foreach ($xml->si as $val) {
            $sharedStrings[] = (string)($val->t ?? $val->r->t ?? '');
        }
    }

    $sheetMap = [];
    if (($wbXml = $zip->getFromName('xl/workbook.xml')) !== false) {
        $xml = simplexml_load_string($wbXml);
        foreach ($xml->sheets->sheet as $s) {
            $sheetMap[(string)$s->attributes('http://schemas.openxmlformats.org/officeDocument/2006/relationships')['id']] = (string)$s['name'];
        }
    }

    $fileMap = [];
    if (($relsXml = $zip->getFromName('xl/_rels/workbook.xml.rels')) !== false) {
        $xml = simplexml_load_string($relsXml);
        foreach ($xml->Relationship as $rel) {
            $id = (string)$rel['Id'];
            $target = (string)$rel['Target'];
            if (isset($sheetMap[$id])) {
                $fileMap[$sheetMap[$id]] = 'xl/' . ltrim($target, '/');
            }
        }
    }

    $allPlayers = [];

    foreach (['P', 'D', 'C', 'A'] as $sheetName) {
        if (!isset($fileMap[$sheetName])) continue;

        $sheetXmlContent = $zip->getFromName($fileMap[$sheetName]);
        if ($sheetXmlContent === false) continue;

        $xml = simplexml_load_string($sheetXmlContent);
        $rowsData = [];

        foreach ($xml->sheetData->row as $r) {
            $rowCells = [];
            foreach ($r->c as $c) {
                $colIdx = excelColumnToIndex((string)$c['r']);
                $t = (string)($c['t'] ?? '');
                $v = (string)($c->v ?? '');
                $rowCells[$colIdx] = ($t === 's' && isset($sharedStrings[(int)$v])) ? $sharedStrings[(int)$v] : $v;
            }
            ksort($rowCells);
            $rowsData[] = $rowCells;
        }

        if (count($rowsData) < 2) continue;

        $header = [];
        foreach ($rowsData[0] as $cIdx => $cVal) {
            $header[$cIdx] = cleanHeaderStr($cVal);
        }

        $idxNome         = array_search('nome', $header);
        $idxRuolo        = array_search('ruolo', $header);
        $idxTeam         = array_search('team', $header);
        $idxFascia       = array_search('fascia', $header);
        $idxPrezzo       = array_search('prezzo', $header);
        $idxPma          = array_search('pma', $header);
        $idxQuo          = array_search('quo', $header);
        $idxTitolarita   = array_search('titolarita', $header);
        $idxAffidabilita = array_search('affidabilita', $header);
        $idxIntegrita    = array_search('integrita', $header);
        $idxMV           = array_search('mv', $header);
        $idxFMV          = array_search('fmv', $header);
        $idxPresenze     = array_search('presenze', $header);
        $idxFmvExp       = array_search('fmvexp', $header);
        $idxPtTit        = array_search('pttit', $header);
        $idxMinuti       = array_search('minuti', $header);
        $idxPtInf        = array_search('ptinf', $header);
        $idxGol          = array_search('gol', $header);
        $idxAssist       = array_search('assist', $header);
        $idxAmm          = array_search('ammonizioni', $header);
        $idxEsp          = array_search('espulsioni', $header);
        $idxRigSeg       = array_search('rigsegnati', $header);
        $idxRigSba       = array_search('rigsbagliati', $header);
        $idxGolSub       = array_search('golsubiti', $header);
        $idxRigPar       = array_search('rigparati', $header);

        $noteCols = [];
        foreach ($header as $cIdx => $cName) {
            if (str_contains($cName, 'nota') || $cName === 'commento') {
                $noteCols[] = $cIdx;
            }
        }

        if ($idxNome === false) continue;

        for ($i = 1; $i < count($rowsData); $i++) {
            $row = $rowsData[$i];
            $nomeVal = trim((string)($row[$idxNome] ?? ''));
            if (empty($nomeVal)) continue;

            $ruoloVal = ($idxRuolo !== false && !empty($row[$idxRuolo])) 
                ? strtoupper(trim((string)$row[$idxRuolo])) 
                : $sheetName;

            if (!in_array($ruoloVal, ['P', 'D', 'C', 'A'])) $ruoloVal = $sheetName;

            $noteArray = [];
            foreach ($noteCols as $nIdx) {
                if (!empty($row[$nIdx])) {
                    $valNote = trim((string)$row[$nIdx]);
                    if ($valNote !== '') $noteArray[] = $valNote;
                }
            }
            $noteStr = implode(' | ', $noteArray);
            $noteLower = strtolower($noteStr);

            $allPlayers[] = [
                'id' => uniqid('p_'),
                'nome' => $nomeVal,
                'squadra' => $idxTeam !== false ? trim((string)($row[$idxTeam] ?? 'ND')) : 'ND',
                'ruolo' => $ruoloVal,
                'fascia' => $idxFascia !== false ? trim((string)($row[$idxFascia] ?? 'ND')) : 'ND',
                'rigorista' => ($ruoloVal !== 'P') && (str_contains($noteLower, 'rigoris') || str_contains($noteLower, 'rigoristi')),
                'pararigori' => ($ruoloVal === 'P') && str_contains($noteLower, 'pararigori'),
                'imbattibilita' => str_contains($noteLower, 'imbattibil'),
                'punizioni' => str_contains($noteLower, 'tiratore') || str_contains($noteLower, 'punizioni'),
                'titolare' => str_contains($noteLower, 'titolarissimo'),
                'preferito' => false,
                'stato' => 'da_comprate',
                'prezzo_acquisto' => 0,
                'budget_max' => $idxPrezzo !== false ? max(1, (int)($row[$idxPrezzo] ?? 1)) : 1,
                'pma' => $idxPma !== false ? trim((string)($row[$idxPma] ?? '-')) : '-',
                'quo' => $idxQuo !== false ? (int)($row[$idxQuo] ?? 0) : 0,
                'titolarita_voto' => $idxTitolarita !== false ? (int)($row[$idxTitolarita] ?? 0) : 0,
                'affidabilita' => $idxAffidabilita !== false ? (int)($row[$idxAffidabilita] ?? 0) : 0,
                'integrita' => $idxIntegrita !== false ? (int)($row[$idxIntegrita] ?? 0) : 0,
                'mv' => $idxMV !== false ? (float)($row[$idxMV] ?? 0.0) : 0.0,
                'fmv' => $idxFMV !== false ? (float)($row[$idxFMV] ?? 0.0) : 0.0,
                'presenze' => $idxPresenze !== false ? (int)($row[$idxPresenze] ?? 0) : 0,
                'fmv_exp' => $idxFmvExp !== false ? (float)($row[$idxFmvExp] ?? 0.0) : 0.0,
                'pt_tit' => $idxPtTit !== false ? (int)($row[$idxPtTit] ?? 0) : 0,
                'minuti' => $idxMinuti !== false ? (int)($row[$idxMinuti] ?? 0) : 0,
                'pt_inf' => $idxPtInf !== false ? (int)($row[$idxPtInf] ?? 0) : 0,
                'gol' => $idxGol !== false ? (int)($row[$idxGol] ?? 0) : 0,
                'assist' => $idxAssist !== false ? (int)($row[$idxAssist] ?? 0) : 0,
                'ammonizioni' => $idxAmm !== false ? (int)($row[$idxAmm] ?? 0) : 0,
                'espulsioni' => $idxEsp !== false ? (int)($row[$idxEsp] ?? 0) : 0,
                'rig_segnati' => $idxRigSeg !== false ? (int)($row[$idxRigSeg] ?? 0) : 0,
                'rig_sbagliati' => $idxRigSba !== false ? (int)($row[$idxRigSba] ?? 0) : 0,
                'gol_subiti' => $idxGolSub !== false ? (int)($row[$idxGolSub] ?? 0) : 0,
                'rig_parati' => $idxRigPar !== false ? (int)($row[$idxRigPar] ?? 0) : 0,
                'note' => $noteStr
            ];
        }
    }

    $zip->close();
    return $allPlayers;
}

// UPLOAD AJAX
if (isset($_GET['action']) && $_GET['action'] === 'upload_file') {
    header('Content-Type: application/json; charset=utf-8');

    if (!isset($_FILES['file_upload'])) {
        echo json_encode(['success' => false, 'message' => 'Nessun file inviato.']);
        exit;
    }

    $file = $_FILES['file_upload'];
    $ext = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));

    if ($ext === 'json') {
        $decoded = json_decode(file_get_contents($file['tmp_name']), true);
        if ($decoded) {
            file_put_contents($dataFile, json_encode($decoded, JSON_PRETTY_PRINT));
            echo json_encode(['success' => true, 'message' => 'Database JSON ripristinato!']);
        } else {
            echo json_encode(['success' => false, 'message' => 'JSON non valido.']);
        }
        exit;
    } elseif (in_array($ext, ['xlsx', 'xls'])) {
        $nuovi = parseXlsxNative($file['tmp_name']);
        if (is_array($nuovi) && isset($nuovi['error'])) {
            echo json_encode(['success' => false, 'message' => $nuovi['message']]);
            exit;
        }
        $data[$currentLega]['giocatori'] = $nuovi;
        file_put_contents($dataFile, json_encode($data, JSON_PRETTY_PRINT));
        echo json_encode(['success' => true, 'message' => 'Importati ' . count($nuovi) . ' calciatori!']);
        exit;
    }
}

// AZIONI POST FORM
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $action = $_POST['action'] ?? '';

    if ($action === 'toggle_favorite') {
        foreach ($data[$currentLega]['giocatori'] as &$p) {
            if ($p['id'] === $_POST['player_id']) { $p['preferito'] = !($p['preferito'] ?? false); break; }
        }
        file_put_contents($dataFile, json_encode($data, JSON_PRETTY_PRINT));
        header("Location: index.php?lega=" . $currentLega);
        exit;
    }

    if ($action === 'clear_database') {
        $data[$currentLega]['giocatori'] = [];
        file_put_contents($dataFile, json_encode($data, JSON_PRETTY_PRINT));
        header("Location: index.php?lega=" . $currentLega);
        exit;
    }

    if ($action === 'update_budget') {
        $data[$currentLega]['budget_iniziale'] = max(1, (int)$_POST['budget_iniziale']);
        $data[$currentLega]['budget_ruoli'] = ['P' => (int)($_POST['b_p'] ?? 0), 'D' => (int)($_POST['b_d'] ?? 0), 'C' => (int)($_POST['b_c'] ?? 0), 'A' => (int)($_POST['b_a'] ?? 0)];
        file_put_contents($dataFile, json_encode($data, JSON_PRETTY_PRINT));
        header("Location: index.php?lega=" . $currentLega);
        exit;
    }

    if ($action === 'update_status') {
        foreach ($data[$currentLega]['giocatori'] as &$p) {
            if ($p['id'] === $_POST['player_id']) {
                $p['stato'] = $_POST['status'];
                $p['prezzo_acquisto'] = ($_POST['status'] === 'preso') ? (int)($_POST['prezzo_acquisto'] ?? 0) : 0;
                break;
            }
        }
        file_put_contents($dataFile, json_encode($data, JSON_PRETTY_PRINT));
        header("Location: index.php?lega=" . $currentLega);
        exit;
    }

    if ($action === 'delete') {
        $data[$currentLega]['giocatori'] = array_values(array_filter($data[$currentLega]['giocatori'], fn($p) => $p['id'] !== $_POST['player_id']));
        file_put_contents($dataFile, json_encode($data, JSON_PRETTY_PRINT));
        header("Location: index.php?lega=" . $currentLega);
        exit;
    }
}

$players = $data[$currentLega]['giocatori'] ?? [];
$budgetIniziale = $data[$currentLega]['budget_iniziale'] ?? 500;
$budgetRuoli = $data[$currentLega]['budget_ruoli'] ?? ['P' => 40, 'D' => 80, 'C' => 150, 'A' => 230];

// CALCOLI FINANZIARI
$spesoTotale = 0;
$spesoRuolo = ['P' => 0, 'D' => 0, 'C' => 0, 'A' => 0];
$slotMax = ['P' => 3, 'D' => 8, 'C' => 8, 'A' => 6];
$slotPresi = ['P' => 0, 'D' => 0, 'C' => 0, 'A' => 0];
$giocatoriCampo = ['P' => [], 'D' => [], 'C' => [], 'A' => []];

foreach ($players as $p) {
    if (($p['stato'] ?? '') === 'preso') {
        $pVal = (int)($p['prezzo_acquisto'] ?? 0);
        $spesoTotale += $pVal;
        $r = $p['ruolo'];
        if (isset($slotPresi[$r])) {
            $slotPresi[$r]++;
            $spesoRuolo[$r] += $pVal;
            $giocatoriCampo[$r][] = $p;
        }
    }
}

$budgetResiduo = $budgetIniziale - $spesoTotale;
$slotMancanti = [
    'P' => max(0, $slotMax['P'] - $slotPresi['P']),
    'D' => max(0, $slotMax['D'] - $slotPresi['D']),
    'C' => max(0, $slotMax['C'] - $slotPresi['C']),
    'A' => max(0, $slotMax['A'] - $slotPresi['A']),
];
$totaleSlotMancanti = array_sum($slotMancanti);
$mediaCreditiPerSlot = $totaleSlotMancanti > 0 ? round($budgetResiduo / $totaleSlotMancanti, 1) : 0;

// CALCOLO SOGLIA DI SALVAGUARDIA CREDITI
$creditiRiservatiSogliaMinima = max(0, $totaleSlotMancanti - 1);
$maxCreditiRilancioAssoluto = max(1, $budgetResiduo - $creditiRiservatiSogliaMinima);
$isZonaRossaBudget = ($budgetResiduo <= $totaleSlotMancanti && $totaleSlotMancanti > 0);

// GESTIONE SMART RE-PAIRING / ALLERTE COPPIE
$aiAdviceList = [];
if (array_sum($slotPresi) === 0) {
    $aiAdviceList[] = ["type" => "info", "msg" => "Asta appena iniziata! Mantieni la calma e focalizzati sui top target."];
}

if ($isZonaRossaBudget) {
    $aiAdviceList[] = ["type" => "danger", "msg" => "🚨 <strong>ALLERTA BUDGET CRITICO:</strong> Hai esattamente {$budgetResiduo} FM per {$totaleSlotMancanti} slot mancanti! Puoi fare solo acquisti a 1 FM."];
}

// Analisi Coppie Interrotte / Giocatori Presi da Avversari
foreach ($giocatoriCampo as $rKey => $listaPresi) {
    foreach ($listaPresi as $pPreso) {
        $sq = $pPreso['squadra'];
        $compagniNonPresi = array_filter($players, fn($x) => $x['squadra'] === $sq && $x['ruolo'] === $rKey && $x['id'] !== $pPreso['id']);
        
        foreach ($compagniNonPresi as $compagno) {
            if ($compagno['stato'] === 'perso') {
                $sostitutiLiberi = array_filter($players, fn($x) => $x['ruolo'] === $rKey && $x['squadra'] !== $sq && $x['stato'] === 'da_comprate');
                usort($sostitutiLiberi, fn($a, $b) => ($b['fmv'] ?? 0) <=> ($a['fmv'] ?? 0));
                $migliorSostituto = $sostitutiLiberi[0]['nome'] ?? 'un nuovo titolare';
                $sqSostituto = $sostitutiLiberi[0]['squadra'] ?? '';

                $aiAdviceList[] = [
                    "type" => "danger",
                    "msg" => "⚠️ <strong>ALLERTA COPPIE:</strong> Hai acquistato <strong>{$pPreso['nome']}</strong> ma la sua riserva <strong>{$compagno['nome']}</strong> è stata presa da un rivale! <br>💡 <em>Consiglio IA:</em> Virare subito su <strong>$migliorSostituto ($sqSostituto)</strong> per completare la rotazione!"
                ];
            }
        }
    }
}

foreach (['P' => 'Portieri', 'D' => 'Difensori', 'C' => 'Centrocampisti', 'A' => 'Attaccanti'] as $rKey => $rLabel) {
    $targetR = $budgetRuoli[$rKey] ?? 0;
    $spesoR = $spesoRuolo[$rKey];
    if ($spesoR > $targetR) {
        $diff = $spesoR - $targetR;
        $aiAdviceList[] = ["type" => "warning", "msg" => "Hai superato il budget target per i <strong>$rLabel</strong> di +$diff FM."];
    }
}

$roleBadges = ['P' => 'bg-warning text-dark', 'D' => 'bg-primary', 'C' => 'bg-info text-dark', 'A' => 'bg-danger'];
?>
<!DOCTYPE html>
<html lang="it" data-bs-theme="dark">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>FantaStrategy LIVE ENGINE ⚡</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
    <style>
        body { background-color: #080e1e; color: #f8fafc; font-family: 'Segoe UI', system-ui, sans-serif; }
        .card-custom { background-color: #121f38; border: 1px solid #1e3256; border-radius: 12px; }
        .badge-role { width: 32px; height: 32px; display: inline-flex; align-items: center; justify-content: center; font-weight: bold; border-radius: 8px; }
        .table-custom { color: #f8fafc; font-size: 0.88rem; }
        .table-custom tbody tr { cursor: pointer; transition: background-color 0.15s ease; }
        .table-custom tbody tr:hover { background-color: #1e3256 !important; }
        .player-preso { background-color: rgba(25, 135, 84, 0.22) !important; }
        .player-perso { opacity: 0.35; filter: grayscale(1); background-color: rgba(220, 53, 69, 0.15) !important; }
        .kpi-card { background: linear-gradient(135deg, #121f38 0%, #080e1e 100%); border: 1px solid #1e3256; border-radius: 10px; }
        .badge-affarone { background-color: #198754; color: #fff; font-weight: bold; }
        .badge-overprice { background-color: #dc3545; color: #fff; font-weight: bold; }
        .soccer-field { background: linear-gradient(180deg, #1e4620 0%, #143016 100%); border: 2px solid #2e6931; border-radius: 12px; padding: 12px; }
        .field-line { border-bottom: 1px dashed rgba(255,255,255,0.2); margin: 8px 0; }
        .field-player-badge { background: rgba(0,0,0,0.7); border: 1px solid #ffc107; color: #fff; padding: 3px 9px; border-radius: 12px; font-size: 0.78rem; display: inline-block; }
        .field-player-badge.riserva { border-color: #6c757d; opacity: 0.85; background: rgba(30, 40, 60, 0.8); }
        .progress-box { width: 100%; background-color: #1a2744; border-radius: 5px; overflow: hidden; display: none; }
        .progress-bar-custom { width: 0%; height: 22px; background-color: #198754; text-align: center; line-height: 22px; color: white; font-size: 12px; font-weight: bold; }
        .note-text { font-size: 0.78rem; color: #9aa8c3; display: block; margin-top: 4px; white-space: pre-line; }
        .btn-fav { cursor: pointer; color: #6c757d; transition: color 0.2s; }
        .btn-fav.active { color: #ffc107; }
        .ai-card { border-left: 4px solid #0dcaf0; background: rgba(13, 202, 240, 0.05); }
        .advisor-box { background: rgba(255, 193, 7, 0.07); border-left: 4px solid #ffc107; border-radius: 6px; padding: 12px; }
        .advisor-player-link { cursor: pointer; text-decoration: underline; font-weight: bold; color: #ffc107 !important; transition: color 0.15s; }
        .advisor-player-link:hover { color: #ffffff !important; }
        .ai-generator-card { background: linear-gradient(135deg, #112347 0%, #0c1830 100%); border: 1px solid #254170; border-radius: 12px; }
        .suggested-player-badge { background: #0d1b33; border: 1px solid #203e68; border-radius: 8px; padding: 8px 12px; }
        .cal-box { background: #081124; border: 1px solid #1a2f57; border-radius: 10px; padding: 10px; }
        .combo-card { background: #0b172e; border: 1px solid #1c3866; border-radius: 10px; transition: transform 0.2s; }
        .combo-card:hover { border-color: #ffc107; transform: translateY(-2px); }
        .tactical-badge { font-size: 0.72rem; padding: 3px 7px; border-radius: 6px; margin-right: 3px; font-weight: bold; }

        /* MODALITÀ ASTA VELOCE UI High-Density */
        body.fast-auction-mode .toggle-hide-fast { display: none !important; }
        body.fast-auction-mode .table-custom td, body.fast-auction-mode .table-custom th { padding: 4px 8px !important; font-size: 0.8rem !important; }
        body.fast-auction-mode .note-text { display: none !important; }

        /* ANIMAZIONE E STILI PROMEMORIA BACKUP INTELLIGENTE */
        @keyframes pulse-backup {
            0% { transform: scale(1); box-shadow: 0 0 0 0 rgba(255, 193, 7, 0.8); }
            70% { transform: scale(1.05); box-shadow: 0 0 0 10px rgba(255, 193, 7, 0); }
            100% { transform: scale(1); box-shadow: 0 0 0 0 rgba(255, 193, 7, 0); }
        }
        .btn-backup-alert {
            animation: pulse-backup 1.5s infinite;
            background-color: #ffc107 !important;
            color: #000 !important;
            font-weight: bold !important;
            border-color: #ffc107 !important;
        }
        #backupTimerBanner {
            background: linear-gradient(90deg, #12213d 0%, #1e3a6e 100%);
            border-bottom: 2px solid #ffc107;
            font-size: 0.85rem;
        }
    </style>
</head>
<body>

<!-- BANNER NOTIFICA BACKUP AUTOMATICO -->
<div id="backupTimerBanner" class="py-1 px-3 text-center text-warning fw-bold d-flex align-items-center justify-content-center gap-2" style="display: none !important;">
    <i class="fa-solid fa-triangle-exclamation text-warning fs-6"></i>
    <span id="backupBannerText">PROMEMORIA ASTA: Sono trascorsi 15 minuti dall'ultimo salvataggio o hai registrato nuovi acquisti. Fai un backup per non rischiare di perdere dati!</span>
    <a href="index.php?lega=<?php echo $currentLega; ?>&action=export_backup&type=json" class="btn btn-xs btn-warning text-dark fw-bold ms-2 px-2 py-0" onclick="resetBackupTimer()"><i class="fa-solid fa-download me-1"></i> Salva Ora</a>
</div>

<!-- NAVBAR CON PULSANTI EXPORT BACKUP, UNDO E MODALITÀ ASTA VELOCE -->
<nav class="navbar navbar-expand-lg navbar-dark bg-dark border-bottom border-secondary sticky-top shadow">
    <div class="container-fluid px-4">
        <a class="navbar-brand fw-bold text-warning fs-4" href="#"><i class="fa-solid fa-crown me-2"></i>FantaStrategy LIVE ENGINE</a>
        <div class="d-flex align-items-center gap-2 flex-wrap">
            <span class="badge bg-secondary"><?php echo count($players); ?> Calciatori</span>
            
            <!-- TASTO ANNULLA ULTIMA AZIONE (UNDO) -->
            <button id="btnUndoAction" class="btn btn-sm btn-outline-warning fw-bold d-none" onclick="eseguiUndo()"><i class="fa-solid fa-rotate-left me-1"></i> Annulla Ultimo Clic</button>

            <!-- TOGGLE VISTA ASTA VELOCE -->
            <button class="btn btn-sm btn-info fw-bold text-dark" onclick="toggleFastMode()"><i class="fa-solid fa-bolt me-1"></i> <span id="fastModeLabel">Vista Veloce</span></button>

            <!-- TASTI DI EXPORT BACKUP -->
            <a id="btnExportJson" href="index.php?lega=<?php echo $currentLega; ?>&action=export_backup&type=json" class="btn btn-sm btn-success fw-bold" onclick="resetBackupTimer()"><i class="fa-solid fa-download me-1"></i> Scarica JSON</a>
            <a href="index.php?lega=<?php echo $currentLega; ?>&action=export_backup&type=csv" class="btn btn-sm btn-outline-success fw-bold"><i class="fa-solid fa-file-csv me-1"></i> Esporta CSV</a>
            
            <button class="btn btn-sm btn-outline-warning" data-bs-toggle="modal" data-bs-target="#importModal"><i class="fa-solid fa-file-excel me-1"></i> Carica JSON/Excel</button>
            <form method="POST" style="display:inline;" onsubmit="return confirm('Sei sicuro di voler svuotare interamente il database dei calciatori?');">
                <input type="hidden" name="action" value="clear_database">
                <button type="submit" class="btn btn-sm btn-outline-danger"><i class="fa-solid fa-trash-can me-1"></i> Svuota DB</button>
            </form>
        </div>
    </div>
</nav>

<div class="container-fluid px-4 py-4">

    <!-- KPI FINANZIARI & AI ADVISOR LIVE -->
    <div class="row g-3 mb-4">
        <div class="col-xl-8 col-lg-7">
            <div class="row g-3 mb-3">
                <div class="col-md-3">
                    <div class="card card-custom p-3 kpi-card">
                        <small class="text-muted text-uppercase fw-bold">Budget Iniziale</small>
                        <h3 class="m-0 text-white"><?php echo $budgetIniziale; ?> <small class="fs-6">FM</small></h3>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card card-custom p-3 kpi-card">
                        <small class="text-muted text-uppercase fw-bold">Speso Totale</small>
                        <h3 class="m-0 text-danger"><?php echo $spesoTotale; ?> <small class="fs-6">FM</small></h3>
                        <small class="text-muted"><?php echo array_sum($slotPresi); ?> / 25 Acquistati</small>
                    </div>
                </div>

                <!-- ALERT ZONA ROSSA CREDITI NEL KPI -->
                <div class="col-md-3">
                    <div class="card card-custom p-3 kpi-card <?php echo $isZonaRossaBudget ? 'border-danger bg-danger text-white' : 'border-success'; ?>">
                        <small class="text-uppercase fw-bold <?php echo $isZonaRossaBudget ? 'text-white' : 'text-muted'; ?>">Budget Residuo</small>
                        <h3 class="m-0 fw-bold <?php echo $isZonaRossaBudget ? 'text-white' : 'text-success'; ?>"><?php echo $budgetResiduo; ?> <small class="fs-6">FM</small></h3>
                        <?php if ($totaleSlotMancanti > 0): ?>
                            <small class="<?php echo $isZonaRossaBudget ? 'text-white fw-bold' : 'text-warning'; ?>">Max 1 Rilancio: <strong><?php echo $maxCreditiRilancioAssoluto; ?> FM</strong></small>
                        <?php endif; ?>
                    </div>
                </div>

                <div class="col-md-3">
                    <div class="card card-custom p-3 kpi-card">
                        <small class="text-muted text-uppercase fw-bold">Media / Slot</small>
                        <h3 class="m-0 text-info"><?php echo $mediaCreditiPerSlot; ?> <small class="fs-6">FM</small></h3>
                        <small class="text-muted"><?php echo $totaleSlotMancanti; ?> slot liberi</small>
                    </div>
                </div>
            </div>

            <div class="card card-custom p-3 ai-card">
                <div class="d-flex align-items-center mb-2">
                    <i class="fa-solid fa-robot text-info fs-4 me-2"></i>
                    <h6 class="m-0 text-info fw-bold">Fanta-Advisor AI Live (Consigli in Tempo Reale)</h6>
                </div>
                <ul class="mb-0 ps-3 small">
                    <?php foreach ($aiAdviceList as $advice): ?>
                        <li class="mb-1 text-light"><?php echo $advice['msg']; ?></li>
                    <?php endforeach; ?>
                </ul>
            </div>
        </div>

        <div class="col-xl-4 col-lg-5">
            <div class="card card-custom p-3 border-warning h-100">
                <small class="text-warning text-uppercase fw-bold mb-1"><i class="fa-solid fa-calculator me-1"></i> Simulatore Rilancio "What-If"</small>
                <div class="input-group input-group-sm mb-2">
                    <span class="input-group-text bg-dark border-secondary text-muted">Offerta IPOTETICA FM</span>
                    <input type="number" id="simBidInput" class="form-control border-secondary text-warning fw-bold" placeholder="Es. 45" min="1" oninput="simulaOfferta()">
                </div>
                <div class="d-flex justify-content-between small">
                    <span class="text-muted">Residuo Simulato: <strong id="simResiduo" class="text-white">-</strong></span>
                    <span class="text-muted">Nuova Media/Slot: <strong id="simMedia" class="text-info">-</strong></span>
                </div>
            </div>
        </div>
    </div>

    <!-- ALGORITMO ABBINAMENTI STRATEGICI -->
    <div class="row g-4 mb-4 toggle-hide-fast">
        <div class="col-12">
            <div class="card card-custom p-3 border-info shadow">
                <div class="d-flex justify-content-between align-items-center flex-wrap gap-2 mb-3 border-bottom border-secondary pb-2">
                    <div class="d-flex align-items-center gap-2">
                        <i class="fa-solid fa-diagram-project text-info fs-3"></i>
                        <div>
                            <h5 class="m-0 text-info fw-bold">Algoritmo Abbinamenti Strategici (3 Portieri / 2 Attaccanti)</h5>
                            <small class="text-muted">Analizza Incroci Calendario, Partite Facili/Medie/Difficili, Statistiche e Categorie dei Giocatori nel DB</small>
                        </div>
                    </div>
                    <button class="btn btn-sm btn-outline-info fw-bold" onclick="calcolaAbbinamentiTop()"><i class="fa-solid fa-arrows-rotate me-1"></i> Calcola Abbinamenti Perfetti</button>
                </div>

                <div class="row g-3">
                    <div class="col-lg-6">
                        <h6 class="text-warning fw-bold mb-2"><i class="fa-solid fa-shield-halved me-1"></i> Top Tridenti Portieri (Incroci Calendario Perfetti)</h6>
                        <div id="keeperCombosBox">
                            <div class="text-center py-3 text-muted small"><i class="fa-solid fa-spinner fa-spin me-1"></i> Calcolo incroci portieri in corso...</div>
                        </div>
                    </div>

                    <div class="col-lg-6">
                        <h6 class="text-danger fw-bold mb-2"><i class="fa-solid fa-fire me-1"></i> Top Coppie d'Attacco a Due (Alternanza e Potenziale Bonus)</h6>
                        <div id="strikerCombosBox">
                            <div class="text-center py-3 text-muted small"><i class="fa-solid fa-spinner fa-spin me-1"></i> Calcolo combinazioni attacco in corso...</div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- GENERATORE SQUADRA IA -->
    <div class="row g-4 mb-4 toggle-hide-fast">
        <div class="col-12">
            <div class="card ai-generator-card p-3 shadow">
                <div class="d-flex justify-content-between align-items-center flex-wrap gap-2 mb-3 border-bottom border-secondary pb-2">
                    <div class="d-flex align-items-center gap-2">
                        <i class="fa-solid fa-wand-magic-sparkles text-warning fs-3"></i>
                        <div>
                            <h5 class="m-0 text-warning fw-bold">Generatore Rosa Completa (P, D, C, A) & Incroci Squadre FantaLab</h5>
                            <small class="text-muted">Analizza la difficoltà e gli incroci reali delle 38 giornate di FantaLab per tutte le Squadre</small>
                        </div>
                    </div>
                    <div class="d-flex gap-2">
                        <button type="button" class="btn btn-warning btn-sm fw-bold shadow-sm" onclick="generaRosaIA(false)"><i class="fa-solid fa-microchip me-1"></i> Genera Formazione Ideale (IA)</button>
                        <button type="button" class="btn btn-outline-info btn-sm fw-bold shadow-sm" onclick="generaRosaIA(true)"><i class="fa-solid fa-rotate me-1"></i> Rigenera Rosa (Variante)</button>
                    </div>
                </div>

                <div id="aiSquadContainer">
                    <div class="text-center py-3 text-muted small">
                        <i class="fa-solid fa-circle-info me-1"></i> Clicca su "Genera Formazione Ideale" per creare la rosa bilanciata sui 4 ruoli con i dati esatti delle Squadre da FantaLab!
                    </div>
                </div>

                <div id="calendarRecapBox" class="mt-3 pt-3 border-top border-secondary" style="display: none;">
                    <h6 class="text-warning fw-bold mb-2"><i class="fa-solid fa-calendar-days me-2"></i>Analisi Incroci Calendario FantaLab per Ruolo (38 Giornate)</h6>
                    <div class="row g-2" id="calendarRecapContent"></div>
                </div>
            </div>
        </div>
    </div>

    <!-- CAMPO TATTICO & BUDGET RUOLI -->
    <div class="row g-4 mb-4 toggle-hide-fast">
        <div class="col-lg-7">
            <div class="card card-custom p-3 h-100">
                <h5 class="text-warning mb-2"><i class="fa-solid fa-futbol me-2"></i>Rosa Acquistata sul Campo</h5>
                <div class="soccer-field text-center">
                    <div class="mb-2">
                        <small class="text-danger fw-bold d-block mb-1">ATTACCO (<?php echo count($giocatoriCampo['A']); ?>/6)</small>
                        <div class="d-flex justify-content-center gap-1 flex-wrap">
                            <?php if (empty($giocatoriCampo['A'])): ?><span class="text-muted small">Nessun attaccante preso</span><?php endif; ?>
                            <?php foreach ($giocatoriCampo['A'] as $pA): ?>
                                <span class="field-player-badge border-danger <?php echo empty($pA['titolare']) ? 'riserva' : ''; ?>" data-bs-toggle="tooltip" title="<?php echo !empty($pA['titolare']) ? 'Titolare' : 'Riserva'; ?>">
                                    <?php echo htmlspecialchars($pA['nome']); ?> <?php echo !empty($pA['titolare']) ? '★' : ''; ?>
                                </span>
                            <?php endforeach; ?>
                        </div>
                    </div>
                    <div class="field-line"></div>
                    <div class="mb-2">
                        <small class="text-info fw-bold d-block mb-1">CENTROCAMPO (<?php echo count($giocatoriCampo['C']); ?>/8)</small>
                        <div class="d-flex justify-content-center gap-1 flex-wrap">
                            <?php if (empty($giocatoriCampo['C'])): ?><span class="text-muted small">Nessun centrocampista preso</span><?php endif; ?>
                            <?php foreach ($giocatoriCampo['C'] as $pC): ?>
                                <span class="field-player-badge border-info <?php echo empty($pC['titolare']) ? 'riserva' : ''; ?>" data-bs-toggle="tooltip" title="<?php echo !empty($pC['titolare']) ? 'Titolare' : 'Riserva'; ?>">
                                    <?php echo htmlspecialchars($pC['nome']); ?> <?php echo !empty($pC['titolare']) ? '★' : ''; ?>
                                </span>
                            <?php endforeach; ?>
                        </div>
                    </div>
                    <div class="field-line"></div>
                    <div class="mb-2">
                        <small class="text-primary fw-bold d-block mb-1">DIFESA (<?php echo count($giocatoriCampo['D']); ?>/8)</small>
                        <div class="d-flex justify-content-center gap-1 flex-wrap">
                            <?php if (empty($giocatoriCampo['D'])): ?><span class="text-muted small">Nessun difensore preso</span><?php endif; ?>
                            <?php foreach ($giocatoriCampo['D'] as $pD): ?>
                                <span class="field-player-badge border-primary <?php echo empty($pD['titolare']) ? 'riserva' : ''; ?>" data-bs-toggle="tooltip" title="<?php echo !empty($pD['titolare']) ? 'Titolare' : 'Riserva'; ?>">
                                    <?php echo htmlspecialchars($pD['nome']); ?> <?php echo !empty($pD['titolare']) ? '★' : ''; ?>
                                </span>
                            <?php endforeach; ?>
                        </div>
                    </div>
                    <div class="field-line"></div>
                    <div>
                        <small class="text-warning fw-bold d-block mb-1">PORTA (<?php echo count($giocatoriCampo['P']); ?>/3)</small>
                        <div class="d-flex justify-content-center gap-1 flex-wrap">
                            <?php if (empty($giocatoriCampo['P'])): ?><span class="text-muted small">Nessun portiere preso</span><?php endif; ?>
                            <?php foreach ($giocatoriCampo['P'] as $pP): ?>
                                <span class="field-player-badge border-warning <?php echo empty($pP['titolare']) ? 'riserva' : ''; ?>" data-bs-toggle="tooltip" title="<?php echo !empty($pP['titolare']) ? 'Titolare' : 'Riserva'; ?>">
                                    <?php echo htmlspecialchars($pP['nome']); ?> <?php echo !empty($pP['titolare']) ? '★' : ''; ?>
                                </span>
                            <?php endforeach; ?>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <div class="col-lg-5">
            <div class="card card-custom p-3 h-100">
                <div class="d-flex justify-content-between align-items-center mb-2">
                    <h5 class="text-warning m-0"><i class="fa-solid fa-pie-chart me-2"></i>Pianificazione Ruoli</h5>
                    <button class="btn btn-sm btn-outline-warning" data-bs-toggle="modal" data-bs-target="#budgetModal"><i class="fa-solid fa-sliders me-1"></i> Configura</button>
                </div>
                <div class="row g-2">
                    <?php foreach (['P' => 'Portieri', 'D' => 'Difensori', 'C' => 'Centrocampisti', 'A' => 'Attaccanti'] as $r => $rNome): 
                        $target = $budgetRuoli[$r] ?? 0;
                        $speso = $spesoRuolo[$r];
                        $rimanente = $target - $speso;
                        $isOver = $speso > $target;
                        $presiR = $slotPresi[$r];
                        $maxR = $slotMax[$r];
                    ?>
                        <div class="col-6">
                            <div class="p-2 border rounded bg-dark text-center">
                                <span class="fw-bold d-block"><?php echo $rNome; ?></span>
                                <span class="badge bg-secondary mb-1"><?php echo $presiR; ?> / <?php echo $maxR; ?> Acquistati</span>
                                <small class="text-muted d-block">Target: <?php echo $target; ?> FM</small>
                                <div class="fw-bold mt-1 <?php echo $isOver ? 'text-danger' : 'text-success'; ?>">
                                    <?php echo $isOver ? 'OVER +' . abs($rimanente) : 'Residuo: ' . $rimanente; ?> FM
                                </div>
                            </div>
                        </div>
                    <?php endforeach; ?>
                </div>
            </div>
        </div>
    </div>

    <!-- MAIN CONTENT: TABELLA FULL-WIDTH CON FILTRI TATTICI E FASCE -->
    <div class="row g-4">
        <div class="col-12">
            <div class="card card-custom p-3">
                <div class="row g-2 mb-3">
                    <div class="col-md-3">
                        <div class="input-group input-group-sm">
                            <span class="input-group-text bg-dark border-secondary"><i class="fa-solid fa-magnifying-glass"></i></span>
                            <input type="text" id="searchInput" class="form-control border-secondary" placeholder="Cerca calciatore o Quick-Bid (es. Barella 42)... (Premi 'F' o '/')">
                        </div>
                    </div>
                    
                    <!-- BARRA FILTRI RAPIDI PER RUOLO, FASCIA E STATO -->
                    <div class="col-md-9 d-flex gap-1 justify-content-md-end flex-wrap">
                        <button class="btn btn-sm btn-outline-light filter-btn active" data-role="all">Tutti</button>
                        <button class="btn btn-sm btn-outline-success filter-btn" data-role="liberi"><i class="fa-solid fa-eye me-1"></i> Solo Liberi</button>
                        <button class="btn btn-sm btn-outline-warning filter-btn" data-role="fav"><i class="fa-solid fa-star text-warning me-1"></i> Preferiti</button>
                        <div class="vr mx-1"></div>
                        <!-- FILTRI FASCIA EXCEL -->
                        <button class="btn btn-sm btn-outline-light filter-btn" data-role="f_top">⭐ Top</button>
                        <button class="btn btn-sm btn-outline-light filter-btn" data-role="f_1">1ª Fascia</button>
                        <button class="btn btn-sm btn-outline-light filter-btn" data-role="f_2">2ª Fascia</button>
                        <button class="btn btn-sm btn-outline-light filter-btn" data-role="f_3">3ª Fascia</button>
                        <div class="vr mx-1"></div>
                        <!-- FILTRI TATTICI -->
                        <button class="btn btn-sm btn-outline-danger filter-btn" data-role="rigoristi">⚽ Rigoristi</button>
                        <button class="btn btn-sm btn-outline-info filter-btn" data-role="piazzati">🎯 Piazzati</button>
                        <button class="btn btn-sm btn-outline-success filter-btn" data-role="valorizzati">📈 Valorizzati</button>
                        <div class="vr mx-1"></div>
                        <!-- FILTRI RUOLO -->
                        <button class="btn btn-sm btn-outline-warning filter-btn" data-role="P">P</button>
                        <button class="btn btn-sm btn-outline-primary filter-btn" data-role="D">D</button>
                        <button class="btn btn-sm btn-outline-info filter-btn" data-role="C">C</button>
                        <button class="btn btn-sm btn-outline-danger filter-btn" data-role="A">A</button>
                    </div>
                </div>

                <div class="table-responsive">
                    <table class="table table-custom align-middle" id="playersTable">
                        <thead class="table-dark">
                            <tr>
                                <th style="width: 30px;">⭐</th>
                                <th>R</th>
                                <th>Calciatore</th>
                                <th>Fascia</th>
                                <th>Statistiche</th>
                                <th>Tag & Tattica</th>
                                <th>Prezzo Excel</th>
                                <th>Prezzo Reale</th>
                                <th>Fanta-Index 📊</th>
                                <th class="text-center">Asta Live</th>
                                <th class="text-end">Azione</th>
                            </tr>
                        </thead>
                        <tbody>
                            <?php if (empty($players)): ?>
                                <tr>
                                    <td colspan="11" class="text-center text-muted py-5">Nessun giocatore salvato. Carica il file Excel per popolare automaticamente la dashboard!</td>
                                </tr>
                            <?php else: ?>
                                <?php foreach ($players as $p): 
                                    $squadra = $p['squadra'];
                                    $tattica = $tatticheSquadre[$squadra] ?? null;

                                    $isRigorista = $tattica && in_array($p['nome'], $tattica['rigoristi']);
                                    $isPiazzato = $tattica && (in_array($p['nome'], $tattica['punizioni']) || in_array($p['nome'], $tattica['angoli']));
                                    $isValorizzato = $tattica && array_key_exists($p['nome'], $tattica['valorizzati']);
                                    $isPenalizzato = $tattica && array_key_exists($p['nome'], $tattica['penalizzati']);
                                    $isNascosto = $tattica && in_array($p['nome'], $tattica['nascosti']);
                                    $isGiovane = $tattica && in_array($p['nome'], $tattica['giovani']);

                                    // COMPAGNI DI REPARTO LIBERI IN SQUADRA
                                    $compagniRepartoLiberi = array_filter($players, fn($x) => $x['squadra'] === $squadra && $x['ruolo'] === $p['ruolo'] && $x['id'] !== $p['id'] && ($x['stato'] ?? '') === 'da_comprate');
                                    $hasCompagniLiberi = count($compagniRepartoLiberi) > 0;

                                    $rowClass = '';
                                    if (($p['stato'] ?? '') === 'preso') $rowClass = 'player-preso';
                                    if (($p['stato'] ?? '') === 'perso') $rowClass = 'player-perso';

                                    $indexBadge = '-';
                                    if (($p['stato'] ?? '') === 'preso' && $p['budget_max'] > 0) {
                                        $risparmio = $p['budget_max'] - $p['prezzo_acquisto'];
                                        if ($risparmio >= 10) {
                                            $indexBadge = '<span class="badge badge-affarone"><i class="fa-solid fa-fire me-1"></i>AFFARONE (+' . $risparmio . ' FM)</span>';
                                        } elseif ($risparmio >= 0) {
                                            $indexBadge = '<span class="badge bg-success">PREZZO OK</span>';
                                        } else {
                                            $indexBadge = '<span class="badge badge-overprice">OVERPRICE (' . $risparmio . ' FM)</span>';
                                        }
                                    }

                                    $notesLower = strtolower($p['note'] ?? '');
                                    $isPortiere = ($p['ruolo'] === 'P');
                                    $isFav = $p['preferito'] ?? false;
                                    
                                    $playerJson = htmlspecialchars(json_encode($p), ENT_QUOTES, 'UTF-8');
                                ?>
                                    <tr class="<?php echo $rowClass; ?>" 
                                        data-role="<?php echo $p['ruolo']; ?>" 
                                        data-fascia="<?php echo strtolower(trim($p['fascia'] ?? '')); ?>"
                                        data-stato="<?php echo $p['stato'] ?? 'da_comprate'; ?>"
                                        data-fav="<?php echo $isFav ? '1' : '0'; ?>"
                                        data-rigorista="<?php echo $isRigorista ? '1' : '0'; ?>"
                                        data-piazzato="<?php echo $isPiazzato ? '1' : '0'; ?>"
                                        data-valorizzato="<?php echo $isValorizzato ? '1' : '0'; ?>"
                                        data-nascosto="<?php echo $isNascosto ? '1' : '0'; ?>"
                                        data-giovane="<?php echo $isGiovane ? '1' : '0'; ?>"
                                        onclick="apriDettagliGiocatore(<?php echo $playerJson; ?>, event)">
                                        <td onclick="event.stopPropagation();">
                                            <form method="POST" style="display:inline;">
                                                <input type="hidden" name="action" value="toggle_favorite">
                                                <input type="hidden" name="player_id" value="<?php echo $p['id']; ?>">
                                                <button type="submit" class="btn btn-link p-0 border-0 btn-fav <?php echo $isFav ? 'active' : ''; ?>" data-bs-toggle="tooltip" title="Aggiungi/Rimuovi dai Preferiti">
                                                    <i class="fa-solid fa-star fs-6"></i>
                                                </button>
                                            </form>
                                        </td>
                                        <td><span class="badge <?php echo $roleBadges[$p['ruolo']]; ?> badge-role"><?php echo $p['ruolo']; ?></span></td>
                                        <td>
                                            <span class="fw-bold fs-6 player-name"><?php echo htmlspecialchars($p['nome']); ?></span><br>
                                            <small class="text-muted player-team"><?php echo htmlspecialchars($p['squadra']); ?> (<?php echo $tattica['allenatore'] ?? 'ND'; ?>)</small>
                                        </td>
                                        <td><span class="badge bg-secondary"><?php echo htmlspecialchars($p['fascia'] ?? 'ND'); ?></span></td>
                                        <td>
                                            <small class="d-block text-info">MV: <?php echo number_format($p['mv'] ?? 0, 2); ?> | FMV: <?php echo number_format($p['fmv'] ?? 0, 2); ?></small>
                                            <small class="text-muted">⚽ <?php echo $p['gol'] ?? 0; ?> | 👟 <?php echo $p['assist'] ?? 0; ?> | Quo: <?php echo $p['quo'] ?? 0; ?></small>
                                        </td>
                                        <td>
                                            <div class="mb-1">
                                                <!-- BADGE COPPIA LIBERA -->
                                                <?php if ($hasCompagniLiberi): ?>
                                                    <span class="badge bg-dark border border-warning text-warning tactical-badge" data-bs-toggle="tooltip" title="<?php echo count($compagniRepartoLiberi); ?> compagni di reparto liberi nel DB">🔗 COPPIA DISPONIBILE</span>
                                                <?php endif; ?>

                                                <!-- BADGE TATTICI -->
                                                <?php if ($isRigorista): ?><span class="badge bg-danger tactical-badge">⚽ RIGORISTA</span><?php endif; ?>
                                                <?php if ($isPiazzato): ?><span class="badge bg-info text-dark tactical-badge">🎯 PIAZZATI</span><?php endif; ?>
                                                <?php if ($isValorizzato): ?><span class="badge bg-success tactical-badge">📈 VALORIZZATO</span><?php endif; ?>
                                                <?php if ($isPenalizzato): ?><span class="badge bg-warning text-dark tactical-badge">⚠️ PENALIZZATO</span><?php endif; ?>
                                                <?php if ($isNascosto): ?><span class="badge bg-secondary tactical-badge">🕵️ NASCOSTO</span><?php endif; ?>
                                                <?php if ($isGiovane): ?><span class="badge bg-primary tactical-badge">👶 GIOVANE</span><?php endif; ?>

                                                <!-- BADGE EXCEL ESSENZIALI -->
                                                <?php if (str_contains($notesLower, 'titolarissimo')): ?><span class="badge bg-dark border border-success text-success me-1" data-bs-toggle="tooltip" title="Titolare inamovibile">🔄 Tit</span><?php endif; ?>
                                                <?php if (str_contains($notesLower, 'modificatore')): ?><span class="badge bg-dark border border-primary text-primary me-1" data-bs-toggle="tooltip" title="Ottimo per Modificatore">🛡️ Mod</span><?php endif; ?>
                                                <?php if ($isPortiere && str_contains($notesLower, 'imbattibil')): ?><span class="badge bg-dark border border-info text-info me-1" data-bs-toggle="tooltip" title="Portiere imbattibilità">🛡️ Imbattibilità</span><?php endif; ?>
                                                <?php if ($isPortiere && str_contains($notesLower, 'pararigori')): ?><span class="badge bg-dark border border-warning text-warning me-1" data-bs-toggle="tooltip" title="Para-rigori">🧤 ParaRig</span><?php endif; ?>
                                                <?php if (str_contains($notesLower, 'rischio infortuni')): ?><span class="badge bg-dark border border-danger text-danger me-1" data-bs-toggle="tooltip" title="Rischio infortuni">🚑 Inf</span><?php endif; ?>
                                                <?php if (str_contains($notesLower, 'scommessa')): ?><span class="badge bg-dark border border-purple text-warning me-1" data-bs-toggle="tooltip" title="Scommessa">🎰 Scommessa</span><?php endif; ?>
                                            </div>

                                            <?php if (!empty($p['note'])): ?>
                                                <span class="note-text"><i class="fa-solid fa-comment-dots me-1 text-warning"></i><?php echo htmlspecialchars($p['note']); ?></span>
                                            <?php endif; ?>
                                        </td>
                                        <td class="fw-bold text-info"><?php echo $p['budget_max'] > 0 ? $p['budget_max'] . ' FM' : '-'; ?></td>
                                        <td class="fw-bold text-success"><?php echo ($p['stato'] ?? '') === 'preso' ? $p['prezzo_acquisto'] . ' FM' : '-'; ?></td>
                                        <td><?php echo $indexBadge; ?></td>
                                        <td class="text-center" onclick="event.stopPropagation();">
                                            <?php if (($p['stato'] ?? '') === 'preso'): ?>
                                                <form method="POST" style="display:inline;" onsubmit="registraAzioneUndo('<?php echo $p['id']; ?>', 'da_comprate')">
                                                    <input type="hidden" name="action" value="update_status">
                                                    <input type="hidden" name="player_id" value="<?php echo $p['id']; ?>">
                                                    <input type="hidden" name="status" value="da_comprate">
                                                    <button type="submit" class="btn btn-sm btn-success fw-bold"><i class="fa-solid fa-check-double me-1"></i> Preso</button>
                                                </form>
                                            <?php else: ?>
                                                <div class="btn-group btn-group-sm">
                                                    <button type="button" class="btn btn-sm btn-outline-success" onclick="prompPreso('<?php echo $p['id']; ?>', '<?php echo htmlspecialchars($p['nome'], ENT_QUOTES); ?>')"><i class="fa-solid fa-check"></i> Preso</button>
                                                    <form method="POST" style="display:inline;" onsubmit="registraAzioneUndo('<?php echo $p['id']; ?>', 'da_comprate')">
                                                        <input type="hidden" name="action" value="update_status">
                                                        <input type="hidden" name="player_id" value="<?php echo $p['id']; ?>">
                                                        <input type="hidden" name="status" value="perso">
                                                        <button type="submit" class="btn btn-sm btn-outline-danger" data-bs-toggle="tooltip" title="Segna come perso / Sconsigliato"><i class="fa-solid fa-trash-can"></i></button>
                                                    </form>
                                                </div>
                                            <?php endif; ?>
                                        </td>
                                        <td class="text-end" onclick="event.stopPropagation();">
                                            <form method="POST" style="display:inline;" onsubmit="return confirm('Eliminare questo calciatore?');">
                                                <input type="hidden" name="action" value="delete">
                                                <input type="hidden" name="player_id" value="<?php echo $p['id']; ?>">
                                                <button type="submit" class="btn btn-sm btn-outline-secondary border-0"><i class="fa-solid fa-xmark text-danger"></i></button>
                                            </form>
                                        </td>
                                    </tr>
                                <?php endforeach; ?>
                            <?php endif; ?>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- MODAL SCHEDA DETTAGLIATA GIOCATORE -->
<div class="modal fade" id="playerDetailModal" tabindex="-1">
    <div class="modal-dialog modal-lg modal-dialog-centered">
        <div class="modal-content card-custom border-warning">
            <div class="modal-header border-secondary">
                <div class="d-flex align-items-center gap-2">
                    <span id="modalPlayerRuolo" class="badge badge-role"></span>
                    <div>
                        <h4 id="modalPlayerNome" class="m-0 text-white fw-bold"></h4>
                        <small id="modalPlayerSquadra" class="text-muted"></small>
                    </div>
                </div>
                <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body">

                <!-- MAX BID ALERT DINAMICO -->
                <div class="alert alert-danger d-flex align-items-center justify-content-between p-2 mb-3 border-danger bg-dark">
                    <div>
                        <i class="fa-solid fa-triangle-exclamation text-danger fs-4 me-2"></i>
                        <span class="fw-bold text-white small">Soglia Max Rilancio Sicuro per questo Giocatore:</span>
                    </div>
                    <span id="modalMaxBidCalculated" class="badge bg-danger fs-6 fw-bold">0 FM</span>
                </div>

                <!-- MODULO COMPAGNI DI REPARTO NELLA STESSA SQUADRA -->
                <div id="modalCompagniBox" class="p-2 border border-warning rounded bg-dark mb-3" style="display:none;">
                    <h6 class="text-warning fw-bold mb-1 small"><i class="fa-solid fa-people-arrows me-1"></i> Compagni di Reparto Disponibili nel DB (<span id="modalSquadraRepartoNome"></span>)</h6>
                    <div id="modalCompagniList" class="d-flex flex-wrap gap-2 pt-1"></div>
                </div>

                <!-- 1. FANTA-ADVICE & ANALISI TATTICA -->
                <div class="advisor-box mb-3">
                    <h6 class="text-warning fw-bold mb-1"><i class="fa-solid fa-lightbulb me-1"></i> Fanta-Advice & Analisi Tattica Allenatore</h6>
                    <div id="modalPlayerAdvice" class="small text-light"></div>
                </div>

                <!-- DIFFICOLTÀ CALENDARIO SQUADRA FANTALAB -->
                <div class="cal-box mb-3">
                    <h6 class="text-warning fw-bold mb-1"><i class="fa-solid fa-calendar-alt me-1"></i> Calendario FantaLab Stagione 2026/2027 (<span id="calSquadraNome"></span>)</h6>
                    <div class="row g-2 text-center" id="modalCalendarBadges"></div>
                </div>

                <!-- BALLOTTAGGI LIVE & GERARCHIE PIAZZATI SQUADRA -->
                <div class="p-2 border border-secondary rounded bg-dark mb-3">
                    <h6 class="text-info fw-bold mb-1 small"><i class="fa-solid fa-scale-balanced me-1"></i> Ballottaggi & Gerarchia Piazzati (<span id="squadraPiazzatiNome"></span>)</h6>
                    <div id="modalBallottaggiBox" class="small text-light mb-2"></div>
                    <div class="row g-2 text-center">
                        <div class="col-md-4"><div class="p-1 bg-black rounded border border-secondary"><small class="text-muted d-block fs-7">Rigoristi</small><strong id="modalRigoristiSquadra" class="text-danger small"></strong></div></div>
                        <div class="col-md-4"><div class="p-1 bg-black rounded border border-secondary"><small class="text-muted d-block fs-7">Punizioni</small><strong id="modalPunizioniSquadra" class="text-info small"></strong></div></div>
                        <div class="col-md-4"><div class="p-1 bg-black rounded border border-secondary"><small class="text-muted d-block fs-7">Angoli</small><strong id="modalAngoliSquadra" class="text-warning small"></strong></div></div>
                    </div>
                </div>

                <!-- INDICI D'ASTA E VALUTAZIONE EXCEL -->
                <h6 class="text-warning fw-bold border-bottom border-secondary pb-1 mb-2"><i class="fa-solid fa-chart-pie me-1"></i> Valutazione Asta & Indici Excel</h6>
                <div class="row g-2 text-center mb-3">
                    <div class="col-md-2 col-4"><div class="p-2 bg-dark rounded border border-secondary"><small class="text-muted d-block">Fascia</small><strong id="modalPlayerFascia" class="text-warning fs-6"></strong></div></div>
                    <div class="col-md-2 col-4"><div class="p-2 bg-dark rounded border border-secondary"><small class="text-muted d-block">Quota</small><strong id="modalPlayerQuo" class="text-info fs-6"></strong></div></div>
                    <div class="col-md-2 col-4"><div class="p-2 bg-dark rounded border border-secondary"><small class="text-muted d-block">PMA %</small><strong id="modalPlayerPma" class="text-light fs-6"></strong></div></div>
                    <div class="col-md-2 col-4"><div class="p-2 bg-dark rounded border border-secondary"><small class="text-muted d-block">Titolarità</small><strong id="modalPlayerTitolarita" class="text-success fs-6"></strong></div></div>
                    <div class="col-md-2 col-4"><div class="p-2 bg-dark rounded border border-secondary"><small class="text-muted d-block">Affidabilità</small><strong id="modalPlayerAffidabilita" class="text-info fs-6"></strong></div></div>
                    <div class="col-md-2 col-4"><div class="p-2 bg-dark rounded border border-secondary"><small class="text-muted d-block">Integrità</small><strong id="modalPlayerIntegrita" class="text-danger fs-6"></strong></div></div>
                </div>

                <!-- STATISTICHE GENERALI -->
                <h6 class="text-warning fw-bold border-bottom border-secondary pb-1 mb-2"><i class="fa-solid fa-list-ol me-1"></i> Statistiche di Rendimento</h6>
                <div class="row g-2 text-center mb-3">
                    <div class="col-md-3 col-6"><div class="p-2 bg-dark rounded border border-secondary"><small class="text-muted d-block">MV / FMV</small><span class="fs-6 fw-bold text-light"><span id="modalPlayerMV"></span> / <span id="modalPlayerFMV" class="text-info"></span></span></div></div>
                    <div class="col-md-3 col-6"><div class="p-2 bg-dark rounded border border-secondary"><small class="text-muted d-block">FMV Attesa</small><strong id="modalPlayerFmvExp" class="text-warning fs-6"></strong></div></div>
                    <div class="col-md-3 col-6"><div class="p-2 bg-dark rounded border border-secondary"><small class="text-muted d-block">Presenze (Tit.)</small><span class="fs-6 fw-bold text-light"><span id="modalPlayerPresenze"></span> (<span id="modalPlayerPtTit" class="text-success"></span>)</span></div></div>
                    <div class="col-md-3 col-6"><div class="p-2 bg-dark rounded border border-secondary"><small class="text-muted d-block">Minuti (Infortuni)</small><span class="fs-6 fw-bold text-light"><span id="modalPlayerMinuti"></span>' (<span id="modalPlayerPtInf" class="text-danger"></span>)</span></div></div>
                </div>

                <!-- BONUS & MALUS DETTAGLIATI -->
                <h6 class="text-warning fw-bold border-bottom border-secondary pb-1 mb-2"><i class="fa-solid fa-futbol me-1"></i> Bonus & Malus Storici</h6>
                <div class="row g-2 text-center mb-3">
                    <div class="col"><div class="p-2 bg-dark rounded border border-secondary"><small class="text-muted d-block">Gol</small><strong id="modalPlayerGol" class="text-success fs-6"></strong></div></div>
                    <div class="col"><div class="p-2 bg-dark rounded border border-secondary"><small class="text-muted d-block">Assist</small><strong id="modalPlayerAssist" class="text-info fs-6"></strong></div></div>
                    <div class="col"><div class="p-2 bg-dark rounded border border-secondary"><small class="text-muted d-block">Amm / Esp</small><span class="fs-6 fw-bold"><span id="modalPlayerAmm" class="text-warning"></span>/<span id="modalPlayerEsp" class="text-danger"></span></span></div></div>
                    <div class="col"><div class="p-2 bg-dark rounded border border-secondary"><small class="text-muted d-block">Rigori (S/Sbag/P)</small><span class="fs-6 fw-bold text-light"><span id="modalPlayerRigSeg" class="text-success"></span>/<span id="modalPlayerRigSba" class="text-danger"></span>/<span id="modalPlayerRigPar" class="text-warning"></span></span></div></div>
                    <div class="col"><div class="p-2 bg-dark rounded border border-secondary"><small class="text-muted d-block">Gol Subiti</small><strong id="modalPlayerGolSub" class="text-danger fs-6"></strong></div></div>
                </div>

                <!-- TAG & NOTE ESTESE -->
                <div class="mb-3">
                    <small class="text-muted d-block mb-1 text-uppercase fw-bold">Tag & Caratteristiche:</small>
                    <div id="modalPlayerTags" class="d-flex flex-wrap gap-1 mb-2"></div>
                </div>

                <div>
                    <small class="text-muted d-block mb-1 text-uppercase fw-bold">Note Tattiche & Commenti Excel:</small>
                    <div id="modalPlayerNote" class="p-3 bg-dark border border-secondary rounded text-light small" style="white-space: pre-line;"></div>
                </div>
            </div>
            <div class="modal-footer border-secondary justify-content-between">
                <a id="btnGoogleSearch" href="#" target="_blank" class="btn btn-sm btn-outline-info"><i class="fa-brands fa-google me-1"></i> Cerca ultime news su Fantacalcio.it</a>
                <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Chiudi</button>
            </div>
        </div>
    </div>
</div>

<!-- MODAL BUDGET -->
<div class="modal fade" id="budgetModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content card-custom">
            <div class="modal-header border-secondary">
                <h5 class="modal-title">Imposta Budget Iniziale e Ruoli</h5>
                <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
            </div>
            <form method="POST">
                <div class="modal-body">
                    <input type="hidden" name="action" value="update_budget">
                    <div class="mb-3">
                        <label class="form-label small text-muted fw-bold">Budget Iniziale Crediti Totali (FM)</label>
                        <input type="number" name="budget_iniziale" class="form-control" value="<?php echo $budgetIniziale; ?>" required min="1">
                    </div>
                    <hr class="border-secondary">
                    <small class="text-warning d-block mb-2 fw-bold">Ripartizione Budget Target per Ruolo (FM):</small>
                    <div class="row g-2">
                        <div class="col-6"><label class="form-label small text-muted">Portieri (P)</label><input type="number" name="b_p" class="form-control form-control-sm" value="<?php echo $budgetRuoli['P'] ?? 40; ?>"></div>
                        <div class="col-6"><label class="form-label small text-muted">Difensori (D)</label><input type="number" name="b_d" class="form-control form-control-sm" value="<?php echo $budgetRuoli['D'] ?? 80; ?>"></div>
                        <div class="col-6"><label class="form-label small text-muted">Centrocampisti (C)</label><input type="number" name="b_c" class="form-control form-control-sm" value="<?php echo $budgetRuoli['C'] ?? 150; ?>"></div>
                        <div class="col-6"><label class="form-label small text-muted">Attaccanti (A)</label><input type="number" name="b_a" class="form-control form-control-sm" value="<?php echo $budgetRuoli['A'] ?? 230; ?>"></div>
                    </div>
                </div>
                <div class="modal-footer border-secondary">
                    <button type="submit" class="btn btn-warning w-100 fw-bold">Salva Pianificazione</button>
                </div>
            </form>
        </div>
    </div>
</div>

<!-- MODAL IMPORT -->
<div class="modal fade" id="importModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content card-custom">
            <div class="modal-header border-secondary">
                <h5 class="modal-title">Carica File Excel (.xlsx) o JSON</h5>
                <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
            </div>
            <form id="ajaxUploadForm">
                <div class="modal-body">
                    <input type="file" id="fileInput" class="form-control mb-3" accept=".xlsx, .xls, .json" required>
                    <div class="progress-box mb-2" id="progressBox"><div class="progress-bar-custom" id="progressBar">0%</div></div>
                    <div id="uploadStatus" class="small fw-bold"></div>
                </div>
                <div class="modal-footer border-secondary">
                    <button type="submit" id="btnUploadSubmit" class="btn btn-warning w-100 fw-bold">Avvia Caricamento</button>
                </div>
            </form>
        </div>
    </div>
</div>

<form id="formPreso" method="POST" style="display:none;" onsubmit="registraAzioneUndo(document.getElementById('presoPlayerId').value, 'da_comprate')">
    <input type="hidden" name="action" value="update_status">
    <input type="hidden" name="player_id" id="presoPlayerId">
    <input type="hidden" name="status" value="preso">
    <input type="hidden" name="prezzo_acquisto" id="presoPrezzo">
</form>

<form id="formUndo" method="POST" style="display:none;">
    <input type="hidden" name="action" value="update_status">
    <input type="hidden" name="player_id" id="undoPlayerId">
    <input type="hidden" name="status" id="undoStatus">
    <input type="hidden" name="prezzo_acquisto" value="0">
</form>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>
const tatticheMap = <?php echo json_encode($tatticheSquadre, JSON_UNESCAPED_UNICODE); ?>;
const allDatabasePlayers = <?php echo json_encode($players, JSON_UNESCAPED_UNICODE); ?>;
const teamCalendarsMap = <?php echo json_encode($teamCalendars, JSON_UNESCAPED_UNICODE); ?>;
const budgetResiduoAttuale = <?php echo $budgetResiduo; ?>;
const slotMancantiAttuali = <?php echo $totaleSlotMancanti; ?>;
const slotMancantiPerRuolo = <?php echo json_encode($slotMancanti, JSON_UNESCAPED_UNICODE); ?>;
const budgetRuoliTarget = <?php echo json_encode($budgetRuoli, JSON_UNESCAPED_UNICODE); ?>;

// GESTIONE UNDO (ANNULLA ULTIMA AZIONE)
function registraAzioneUndo(playerId, oldStatus) {
    sessionStorage.setItem('fanta_undo_player_id', playerId);
    sessionStorage.setItem('fanta_undo_status', oldStatus);
}

function controllaPulsanteUndo() {
    const uId = sessionStorage.getItem('fanta_undo_player_id');
    const uBtn = document.getElementById('btnUndoAction');
    if (uId && uBtn) {
        uBtn.classList.remove('d-none');
    }
}

function eseguiUndo() {
    const uId = sessionStorage.getItem('fanta_undo_player_id');
    const uStatus = sessionStorage.getItem('fanta_undo_status') || 'da_comprate';
    if (uId) {
        document.getElementById('undoPlayerId').value = uId;
        document.getElementById('undoStatus').value = uStatus;
        sessionStorage.removeItem('fanta_undo_player_id');
        sessionStorage.removeItem('fanta_undo_status');
        document.getElementById('formUndo').submit();
    }
}

// TOGGLE ASTA VELOCE UI
function toggleFastMode() {
    document.body.classList.toggle('fast-auction-mode');
    const isFast = document.body.classList.contains('fast-auction-mode');
    localStorage.setItem('fanta_fast_mode', isFast ? '1' : '0');
    document.getElementById('fastModeLabel').textContent = isFast ? 'Vista Completa' : 'Vista Veloce';
}

// GESTIONE PROMEMORIA BACKUP VISIVO INTELLIGENTE
let lastBackupTimestamp = localStorage.getItem('fanta_last_backup_time') || Date.now();
const BACKUP_INTERVAL_MINUTES = 15;

function checkBackupTimer() {
    const now = Date.now();
    const elapsedMinutes = Math.floor((now - lastBackupTimestamp) / 60000);
    const banner = document.getElementById('backupTimerBanner');
    const jsonBtn = document.getElementById('btnExportJson');

    if (elapsedMinutes >= BACKUP_INTERVAL_MINUTES) {
        if (banner) banner.style.setProperty('display', 'flex', 'important');
        if (jsonBtn) jsonBtn.classList.add('btn-backup-alert');
    } else {
        if (banner) banner.style.setProperty('display', 'none', 'important');
        if (jsonBtn) jsonBtn.classList.remove('btn-backup-alert');
    }
}

function resetBackupTimer() {
    lastBackupTimestamp = Date.now();
    localStorage.setItem('fanta_last_backup_time', lastBackupTimestamp);
    checkBackupTimer();
}

function getTeamStats(teamCode) {
    const cal = teamCalendarsMap[teamCode] || Array(38).fill(2);
    let f = 0, m = 0, d = 0;
    cal.forEach(v => {
        if (v === 1) f++;
        else if (v === 2) m++;
        else if (v === 3) d++;
    });
    return { facili: f, medie: m, difficili: d };
}

// ALGORITMO ABBINAMENTI PORTIERI E ATTACCANTI
function calcolaAbbinamentiTop() {
    const kBox = document.getElementById('keeperCombosBox');
    const sBox = document.getElementById('strikerCombosBox');

    const portieri = allDatabasePlayers.filter(p => p.ruolo === 'P' && p.stato !== 'perso');
    const attaccanti = allDatabasePlayers.filter(p => p.ruolo === 'A' && p.stato !== 'perso');

    let keeperTeams = [...new Set(portieri.map(p => p.squadra))];
    let keeperCombos = [];

    for (let i = 0; i < keeperTeams.length; i++) {
        for (let j = i + 1; j < keeperTeams.length; j++) {
            for (let k = j + 1; k < keeperTeams.length; k++) {
                const t1 = keeperTeams[i], t2 = keeperTeams[j], t3 = keeperTeams[k];
                const cal1 = teamCalendarsMap[t1] || Array(38).fill(2);
                const cal2 = teamCalendarsMap[t2] || Array(38).fill(2);
                const cal3 = teamCalendarsMap[t3] || Array(38).fill(2);

                let facili = 0, medie = 0, difficili = 0;
                for (let g = 0; g < 38; g++) {
                    const minDiff = Math.min(cal1[g], cal2[g], cal3[g]);
                    if (minDiff === 1) facili++;
                    else if (minDiff === 2) medie++;
                    else difficili++;
                }

                keeperCombos.push({ teams: [t1, t2, t3], facili, medie, difficili, score: (facili * 3) + medie - (difficili * 2) });
            }
        }
    }

    keeperCombos.sort((a, b) => b.score - a.score);
    const topKeeper = keeperCombos.slice(0, 3);

    let kHtml = '';
    topKeeper.forEach((c, idx) => {
        kHtml += `
            <div class="combo-card p-2 mb-2">
                <small class="badge bg-secondary">Opzione #${idx + 1}</small>
                <div class="small text-white my-1"><strong>${c.teams.join(' + ')}</strong></div>
                <div class="d-flex gap-1 text-center small">
                    <span class="badge bg-success flex-fill">🟢 ${c.facili} Facili</span>
                    <span class="badge bg-warning text-dark flex-fill">🟡 ${c.medie} Medie</span>
                    <span class="badge bg-danger flex-fill">🔴 ${c.difficili} Big</span>
                </div>
            </div>
        `;
    });
    kBox.innerHTML = kHtml || '<div class="small text-muted">Carica portieri per vedere gli incroci.</div>';

    let strikerCombos = [];
    let topStrikers = attaccanti.sort((a, b) => parseFloat(b.fmv || 0) - parseFloat(a.fmv || 0)).slice(0, 15);

    for (let i = 0; i < topStrikers.length; i++) {
        for (let j = i + 1; j < topStrikers.length; j++) {
            const a1 = topStrikers[i], a2 = topStrikers[j];
            if (a1.squadra === a2.squadra) continue;

            const cal1 = teamCalendarsMap[a1.squadra] || Array(38).fill(2);
            const cal2 = teamCalendarsMap[a2.squadra] || Array(38).fill(2);

            let ok = 0;
            for (let g = 0; g < 38; g++) { if (cal1[g] === 1 || cal2[g] === 1) ok++; }

            strikerCombos.push({ a1, a2, ok, score: ok });
        }
    }

    strikerCombos.sort((a, b) => b.score - a.score);
    const topStriker = strikerCombos.slice(0, 3);

    let sHtml = '';
    topStriker.forEach((c, idx) => {
        sHtml += `
            <div class="combo-card p-2 mb-2">
                <small class="badge bg-secondary">Coppia #${idx + 1}</small>
                <div class="small text-white my-1"><strong>${c.a1.nome}</strong> (${c.a1.squadra}) + <strong>${c.a2.nome}</strong> (${c.a2.squadra})</div>
                <span class="badge bg-success py-1">⚽ Alternanza Bonus in ${c.ok}/38 Giornate</span>
            </div>
        `;
    });
    sBox.innerHTML = sHtml || '<div class="small text-muted">Carica attaccanti per vedere le coppie.</div>';
}

// GENERATORE IA COMPLETO BILANCIATO
function generaRosaIA(isShuffle = false) {
    const container = document.getElementById('aiSquadContainer');
    const recapBox = document.getElementById('calendarRecapBox');
    
    const disponibili = allDatabasePlayers.filter(p => p.stato !== 'preso' && p.stato !== 'perso');

    if (disponibili.length === 0) {
        container.innerHTML = `<div class="alert alert-warning mb-0 small">Nessun calciatore disponibile nel database per generare la rosa.</div>`;
        return;
    }

    let formazioneConsigliata = { 'P': [], 'D': [], 'C': [], 'A': [] };
    let costoPrevistoTotale = 0;

    let budgetPonderato = {};
    let totaleTargetMancante = 0;

    ['P', 'D', 'C', 'A'].forEach(r => {
        if ((slotMancantiPerRuolo[r] || 0) > 0) {
            let targetBase = budgetRuoliTarget[r] || 50;
            budgetPonderato[r] = targetBase;
            totaleTargetMancante += targetBase;
        } else {
            budgetPonderato[r] = 0;
        }
    });

    let maxBudgetPerRuolo = {};
    ['P', 'D', 'C', 'A'].forEach(r => {
        if (totaleTargetMancante > 0) {
            maxBudgetPerRuolo[r] = Math.round((budgetPonderato[r] / totaleTargetMancante) * budgetResiduoAttuale);
        } else {
            maxBudgetPerRuolo[r] = 0;
        }
    });

    ['P', 'D', 'C', 'A'].forEach(r => {
        const slotsNeeded = slotMancantiPerRuolo[r] || 0;
        if (slotsNeeded <= 0) return;

        let rolePool = disponibili.filter(p => p.ruolo === r).sort((a, b) => {
            const fmvA = parseFloat(a.fmv || a.mv || 0);
            const fmvB = parseFloat(b.fmv || b.mv || 0);
            const calA = getTeamStats(a.squadra).facili;
            const calB = getTeamStats(b.squadra).facili;

            let scoreA = fmvA + (calA * 0.05);
            let scoreB = fmvB + (calB * 0.05);

            if (isShuffle) {
                scoreA += (Math.random() * 3.5);
                scoreB += (Math.random() * 3.5);
            }
            return scoreB - scoreA;
        });

        let spesoNelRuolo = 0;
        const maxSpesaRuolo = maxBudgetPerRuolo[r];

        for (let i = 0; i < rolePool.length; i++) {
            if (formazioneConsigliata[r].length >= slotsNeeded) break;
            const p = rolePool[i];
            const realCost = Math.max(1, parseInt(p.budget_max || 1));
            const slotRimastiNelRuolo = slotsNeeded - formazioneConsigliata[r].length;

            if ((spesoNelRuolo + realCost + (slotRimastiNelRuolo - 1)) <= maxSpesaRuolo && 
                (costoPrevistoTotale + realCost) <= budgetResiduoAttuale) {
                formazioneConsigliata[r].push({ player: p, cost: realCost });
                spesoNelRuolo += realCost;
                costoPrevistoTotale += realCost;
            }
        }

        if (formazioneConsigliata[r].length < slotsNeeded) {
            const presiIds = formazioneConsigliata[r].map(x => x.player.id);
            const rimanentiEconomici = rolePool.filter(p => !presiIds.includes(p.id)).sort((a,b) => (a.budget_max||1) - (b.budget_max||1));

            for (let k = 0; k < rimanentiEconomici.length; k++) {
                if (formazioneConsigliata[r].length >= slotsNeeded) break;
                const p = rimanentiEconomici[k];
                const realCost = Math.max(1, parseInt(p.budget_max || 1));
                formazioneConsigliata[r].push({ player: p, cost: realCost });
                costoPrevistoTotale += realCost;
            }
        }
    });

    let html = `
        <div class="row g-2 mb-2">
            <div class="col-md-6">
                <span class="small text-info fw-bold"><i class="fa-solid fa-calculator me-1"></i> Costo Totale Reale Excel: ${costoPrevistoTotale} FM / ${budgetResiduoAttuale} FM</span>
            </div>
            <div class="col-md-6 text-md-end">
                <span class="small text-success fw-bold"><i class="fa-solid fa-wallet me-1"></i> Crediti Residui in Cassa: ${Math.max(0, budgetResiduoAttuale - costoPrevistoTotale)} FM</span>
            </div>
        </div>
        <div class="row g-2">
    `;

    const labelsRuolo = { 'P': 'Portieri', 'D': 'Difensori', 'C': 'Centrocampisti', 'A': 'Attaccanti' };
    const badgeBg = { 'P': 'bg-warning text-dark', 'D': 'bg-primary', 'C': 'bg-info text-dark', 'A': 'bg-danger' };

    ['A', 'C', 'D', 'P'].forEach(r => {
        const list = formazioneConsigliata[r];
        if (list.length > 0) {
            html += `
                <div class="col-md-3 col-sm-6">
                    <div class="p-2 border border-secondary rounded bg-dark h-100">
                        <small class="fw-bold d-block mb-1 text-uppercase text-warning border-bottom border-secondary pb-1">${labelsRuolo[r]} (${list.length})</small>
            `;
            list.forEach(item => {
                const p = item.player;
                html += `
                    <div class="suggested-player-badge mb-1 d-flex justify-content-between align-items-center">
                        <div>
                            <span class="badge ${badgeBg[r]} me-1" style="font-size: 0.7rem;">${p.ruolo}</span>
                            <strong class="advisor-player-link text-white" style="font-size: 0.85rem;" onclick="caricaSchedaDaNome('${p.nome.replace(/'/g, "\\'")}')">${p.nome}</strong> 
                            <small class="text-muted">(${p.squadra})</small>
                        </div>
                        <div>
                            <span class="badge bg-success me-1">${item.cost} FM</span>
                            <button class="btn btn-xs btn-outline-success p-0 px-1" onclick="prompPreso('${p.id}', '${p.nome.replace(/'/g, "\\'")}')"><i class="fa-solid fa-check"></i></button>
                        </div>
                    </div>
                `;
            });
            html += `</div></div>`;
        }
    });

    html += `</div>`;
    container.innerHTML = html;
    recapBox.style.display = 'block';
}

function caricaSchedaDaNome(nomeGiocatore) {
    const targetPlayer = allDatabasePlayers.find(p => p.nome.toLowerCase() === nomeGiocatore.toLowerCase());
    if (targetPlayer) {
        apriDettagliGiocatore(targetPlayer, { target: document.createElement('div') });
    }
}

function apriDettagliGiocatore(player, event) {
    if (event.target.closest('button') || event.target.closest('form') || event.target.closest('a')) return;

    const modalEl = document.getElementById('playerDetailModal');
    let modal = bootstrap.Modal.getInstance(modalEl) || new bootstrap.Modal(modalEl);

    document.getElementById('modalPlayerNome').textContent = player.nome;
    document.getElementById('modalPlayerSquadra').textContent = player.squadra;
    document.getElementById('squadraPiazzatiNome').textContent = player.squadra;
    
    const altriSlotMancanti = Math.max(0, slotMancantiAttuali - 1);
    const maxBidDisponibile = Math.max(1, budgetResiduoAttuale - altriSlotMancanti);
    document.getElementById('modalMaxBidCalculated').textContent = maxBidDisponibile + " FM";

    // POPOLAMENTO COMPAGNI DI REPARTO NELLA MODAL
    const compagni = allDatabasePlayers.filter(x => x.squadra === player.squadra && x.ruolo === player.ruolo && x.id !== player.id && (x.stato || '') === 'da_comprate');
    const compagniBox = document.getElementById('modalCompagniBox');
    const compagniList = document.getElementById('modalCompagniList');
    
    if (compagni.length > 0) {
        document.getElementById('modalSquadraRepartoNome').textContent = player.squadra + " - " + player.ruolo;
        let cHtml = '';
        compagni.forEach(c => {
            cHtml += `<span class="badge bg-secondary advisor-player-link p-2" onclick="caricaSchedaDaNome('${c.nome.replace(/'/g, "\\'")}')">👤 ${c.nome} (Excel: ${c.budget_max || 1} FM)</span>`;
        });
        compagniList.innerHTML = cHtml;
        compagniBox.style.display = 'block';
    } else {
        compagniBox.style.display = 'none';
    }

    const roleBadge = document.getElementById('modalPlayerRuolo');
    roleBadge.textContent = player.ruolo;
    roleBadge.className = 'badge badge-role ' + (
        player.ruolo === 'P' ? 'bg-warning text-dark' :
        player.ruolo === 'D' ? 'bg-primary' :
        player.ruolo === 'C' ? 'bg-info text-dark' : 'bg-danger'
    );

    document.getElementById('calSquadraNome').textContent = player.squadra;
    const cal = getTeamStats(player.squadra);
    document.getElementById('modalCalendarBadges').innerHTML = `
        <div class="col-4"><span class="badge bg-success w-100 py-2 fs-6">🟢 ${cal.facili} Partite Facili</span></div>
        <div class="col-4"><span class="badge bg-warning text-dark w-100 py-2 fs-6">🟡 ${cal.medie} Partite Medie</span></div>
        <div class="col-4"><span class="badge bg-danger w-100 py-2 fs-6">🔴 ${cal.difficili} Big Match</span></div>
    `;

    document.getElementById('modalPlayerFascia').textContent = player.fascia || 'ND';
    document.getElementById('modalPlayerQuo').textContent = player.quo || 0;
    document.getElementById('modalPlayerPma').textContent = player.pma || '-';
    document.getElementById('modalPlayerTitolarita').textContent = player.titolarita_voto || 0;
    document.getElementById('modalPlayerAffidabilita').textContent = player.affidabilita || 0;
    document.getElementById('modalPlayerIntegrita').textContent = player.integrita || 0;

    document.getElementById('modalPlayerMV').textContent = player.mv ? parseFloat(player.mv).toFixed(2) : '0.00';
    document.getElementById('modalPlayerFMV').textContent = player.fmv ? parseFloat(player.fmv).toFixed(2) : '0.00';
    document.getElementById('modalPlayerFmvExp').textContent = player.fmv_exp ? parseFloat(player.fmv_exp).toFixed(2) : '0.00';
    document.getElementById('modalPlayerPresenze').textContent = player.presenze || 0;
    document.getElementById('modalPlayerPtTit').textContent = player.pt_tit || 0;
    document.getElementById('modalPlayerMinuti').textContent = player.minuti || 0;
    document.getElementById('modalPlayerPtInf').textContent = player.pt_inf || 0;

    document.getElementById('modalPlayerGol').textContent = player.gol || 0;
    document.getElementById('modalPlayerAssist').textContent = player.assist || 0;
    document.getElementById('modalPlayerAmm').textContent = player.ammonizioni || 0;
    document.getElementById('modalPlayerEsp').textContent = player.espulsioni || 0;
    document.getElementById('modalPlayerRigSeg').textContent = player.rig_segnati || 0;
    document.getElementById('modalPlayerRigSba').textContent = player.rig_sbagliati || 0;
    document.getElementById('modalPlayerRigPar').textContent = player.rig_parati || 0;
    document.getElementById('modalPlayerGolSub').textContent = player.gol_subiti || 0;

    const tattica = tatticheMap[player.squadra];
    let adviceHtml = '';

    if (tattica) {
        adviceHtml += `<strong>Mister: ${tattica.allenatore}</strong><br>`;
        if (tattica.valorizzati && tattica.valorizzati[player.nome]) {
            adviceHtml += `<span class="text-success fw-bold">📈 VALORIZZATO DAL MISTER:</span> ${tattica.valorizzati[player.nome]}<br>`;
        } else if (tattica.penalizzati && tattica.penalizzati[player.nome]) {
            adviceHtml += `<span class="text-warning fw-bold">⚠️ PENALIZZATO DAL MISTER:</span> ${tattica.penalizzati[player.nome]}<br>`;
        }

        document.getElementById('modalRigoristiSquadra').textContent = tattica.rigoristi ? tattica.rigoristi.join(', ') : '-';
        document.getElementById('modalPunizioniSquadra').textContent = tattica.punizioni ? tattica.punizioni.join(', ') : '-';
        document.getElementById('modalAngoliSquadra').textContent = tattica.angoli ? tattica.angoli.join(', ') : '-';
        document.getElementById('modalBallottaggiBox').innerHTML = tattica.ballottaggi ? tattica.ballottaggi.map(b => `• ${b}`).join('<br>') : 'Nessun ballottaggio critico.';
    }

    document.getElementById('modalPlayerAdvice').innerHTML = adviceHtml;
    document.getElementById('modalPlayerNote').textContent = player.note || 'Nessuna nota aggiuntiva.';
    document.getElementById('btnGoogleSearch').href = 'https://www.google.com/search?q=' + encodeURIComponent('consigli fantacalcio ' + player.nome + ' ' + player.squadra);

    modal.show();
}

function simulaOfferta() {
    const bid = parseInt(document.getElementById('simBidInput').value) || 0;
    if (bid <= 0) return;
    const nuovoResiduo = budgetResiduoAttuale - bid;
    const nuoviSlot = Math.max(0, slotMancantiAttuali - 1);
    document.getElementById('simResiduo').textContent = nuovoResiduo + " FM";
    document.getElementById('simMedia').textContent = (nuoviSlot > 0 ? (nuovoResiduo / nuoviSlot).toFixed(1) : 0) + " FM";
}

function prompPreso(id, nome) {
    let prezzo = prompt("A quanto hai acquistato " + nome + "? (Crediti FM):", "1");
    if (prezzo !== null && prezzo !== "") {
        registraAzioneUndo(id, 'da_comprate');
        document.getElementById('presoPlayerId').value = id;
        document.getElementById('presoPrezzo').value = parseInt(prezzo) || 0;
        document.getElementById('formPreso').submit();
    }
}

document.addEventListener('DOMContentLoaded', function() {
    calcolaAbbinamentiTop();
    checkBackupTimer();
    controllaPulsanteUndo();
    setInterval(checkBackupTimer, 30000);

    // Ripristina Fast Mode se salvata
    if (localStorage.getItem('fanta_fast_mode') === '1') {
        document.body.classList.add('fast-auction-mode');
        document.getElementById('fastModeLabel').textContent = 'Vista Completa';
    }

    const searchInput = document.getElementById('searchInput');
    const filterBtns = document.querySelectorAll('.filter-btn');
    const rows = document.querySelectorAll('#playersTable tbody tr');
    let currentRole = 'all';

    document.addEventListener('keydown', function(e) {
        if ((e.key === 'f' || e.key === 'F' || e.key === '/') && document.activeElement !== searchInput && document.activeElement.tagName !== 'INPUT') {
            e.preventDefault();
            searchInput.focus();
            searchInput.select();
        }
    });

    function filterTable() {
        const val = searchInput.value.trim();
        const parts = val.split(' ');
        const lastPart = parts[parts.length - 1];
        const isQuickBid = !isNaN(lastPart) && lastPart !== '' && parts.length > 1;
        const term = isQuickBid ? parts.slice(0, -1).join(' ').toLowerCase() : val.toLowerCase();

        rows.forEach(row => {
            const text = row.textContent.toLowerCase();
            const role = row.getAttribute('data-role');
            const fascia = row.getAttribute('data-fascia');
            const stato = row.getAttribute('data-stato');
            const isFav = row.getAttribute('data-fav') === '1';
            const isRigorista = row.getAttribute('data-rigorista') === '1';
            const isPiazzato = row.getAttribute('data-piazzato') === '1';
            const isValorizzato = row.getAttribute('data-valorizzato') === '1';

            const matchSearch = text.includes(term);
            let matchRole = false;

            if (currentRole === 'all') matchRole = true;
            else if (currentRole === 'liberi') matchRole = (stato === 'da_comprate');
            else if (currentRole === 'fav') matchRole = isFav;
            else if (currentRole === 'f_top') matchRole = fascia.includes('top');
            else if (currentRole === 'f_1') matchRole = fascia.includes('1') || fascia.includes('prima');
            else if (currentRole === 'f_2') matchRole = fascia.includes('2') || fascia.includes('seconda');
            else if (currentRole === 'f_3') matchRole = fascia.includes('3') || fascia.includes('terza');
            else if (currentRole === 'rigoristi') matchRole = isRigorista;
            else if (currentRole === 'piazzati') matchRole = isPiazzato;
            else if (currentRole === 'valorizzati') matchRole = isValorizzato;
            else matchRole = (role === currentRole);

            row.style.display = (matchSearch && matchRole) ? '' : 'none';
        });
    }

    searchInput.addEventListener('keydown', function(e) {
        if (e.key === 'Enter') {
            const val = searchInput.value.trim();
            const parts = val.split(' ');
            const lastPart = parts[parts.length - 1];

            if (!isNaN(lastPart) && lastPart !== '' && parts.length > 1) {
                const searchName = parts.slice(0, -1).join(' ').toLowerCase();
                const prezzo = parseInt(lastPart);

                const matchedPlayer = allDatabasePlayers.find(p => p.nome.toLowerCase().includes(searchName) && p.stato !== 'preso');
                if (matchedPlayer) {
                    if (confirm(`Confermi l'acquisto QUICK-BID di ${matchedPlayer.nome} a ${prezzo} FM?`)) {
                        registraAzioneUndo(matchedPlayer.id, 'da_comprate');
                        document.getElementById('presoPlayerId').value = matchedPlayer.id;
                        document.getElementById('presoPrezzo').value = prezzo;
                        document.getElementById('formPreso').submit();
                    }
                } else {
                    alert('Nessun calciatore disponibile trovato con questo nome!');
                }
            }
        }
    });

    searchInput.addEventListener('input', filterTable);

    filterBtns.forEach(btn => {
        btn.addEventListener('click', function() {
            filterBtns.forEach(b => b.classList.remove('active'));
            this.classList.add('active');
            currentRole = this.getAttribute('data-role');
            filterTable();
        });
    });

    document.getElementById('ajaxUploadForm').addEventListener('submit', function(e) {
        e.preventDefault();
        const file = document.getElementById('fileInput').files[0];
        if (!file) return;

        const formData = new FormData();
        formData.append('file_upload', file);

        const xhr = new XMLHttpRequest();
        document.getElementById('progressBox').style.display = 'block';

        xhr.upload.addEventListener('progress', function(e) {
            if (e.lengthComputable) {
                const percent = Math.round((e.loaded / e.total) * 100);
                document.getElementById('progressBar').style.width = percent + '%';
                document.getElementById('progressBar').textContent = percent + '%';
            }
        });

        xhr.addEventListener('load', function() {
            try {
                const res = JSON.parse(xhr.responseText);
                if (res.success) {
                    alert(res.message);
                    window.location.reload();
                } else alert(res.message);
            } catch (err) { alert('Errore nell\'upload'); }
        });

        xhr.open('POST', 'index.php?lega=<?php echo $currentLega; ?>&action=upload_file', true);
        xhr.send(formData);
    });
});
</script>
</body>
</html>