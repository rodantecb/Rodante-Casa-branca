<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$pdo = new PDO("mysql:host=localhost;dbname=relatorios;charset=utf8", "seu_usuario_mysql", "sua_senha_mysql");
$stmt = $pdo->query("SELECT * FROM relatorio_rodante ORDER BY criado_em DESC");
$relatorios = $stmt->fetchAll(PDO::FETCH_ASSOC);

// Criar planilha
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

$sheet->setCellValue('A1', 'ID');
$sheet->setCellValue('B1', 'Técnico');
$sheet->setCellValue('C1', 'Data');
$sheet->setCellValue('D1', 'Dia');
$sheet->setCellValue('E1', 'Observação');
$sheet->setCellValue('F1', 'Linhas');

$row = 2;
foreach ($relatorios as $rel) {
    $linhas = json_decode($rel['linhas'], true);
    $linhasTexto = "";
    foreach($linhas as $linha) {
        $linhasTexto .= "Início: ".$linha['inicio']." - Término: ".$linha['termino']." - Serviço: ".$linha['servico']."\n";
    }

    $sheet->setCellValue("A$row", $rel['id']);
    $sheet->setCellValue("B$row", $rel['tecnico']);
    $sheet->setCellValue("C$row", $rel['data_relatorio']);
    $sheet->setCellValue("D$row", $rel['dia_semana']);
    $sheet->setCellValue("E$row", $rel['observacao']);
    $sheet->setCellValue("F$row", $linhasTexto);
    $row++;
}

// Enviar arquivo para download
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="relatorios_rodante.xlsx"');
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;

<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

// Buscar dados do banco
$pdo = new PDO("mysql:host=localhost;dbname=relatorios;charset=utf8", "seu_usuario_mysql", "sua_senha_mysql");
$stmt = $pdo->query("SELECT * FROM relatorio_rodante ORDER BY criado_em DESC");
$relatorios = $stmt->fetchAll(PDO::FETCH_ASSOC);

// Gerar Excel temporário
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'ID');
$sheet->setCellValue('B1', 'Técnico');
$sheet->setCellValue('C1', 'Data');
$sheet->setCellValue('D1', 'Dia');
$sheet->setCellValue('E1', 'Observação');
$sheet->setCellValue('F1', 'Linhas');

$row = 2;
foreach ($relatorios as $rel) {
    $linhas = json_decode($rel['linhas'], true);
    $linhasTexto = "";
    foreach($linhas as $linha) {
        $linhasTexto .= "Início: ".$linha['inicio']." - Término: ".$linha['termino']." - Serviço: ".$linha['servico']."\n";
    }
    $sheet->setCellValue("A$row", $rel['id']);
    $sheet->setCellValue("B$row", $rel['tecnico']);
    $sheet->setCellValue("C$row", $rel['data_relatorio']);
    $sheet->setCellValue("D$row", $rel['dia_semana']);
    $sheet->setCellValue("E$row", $rel['observacao']);
    $sheet->setCellValue("F$row", $linhasTexto);
    $row++;
}

$arquivo = sys_get_temp_dir() . '/relatorios_envio.xlsx';
$writer = new Xlsx($spreadsheet);
$writer->save($arquivo);

// Enviar por e-mail
$mail = new PHPMailer(true);
try {
    $mail->isSMTP();
    $mail->Host = "smtp.gmail.com";
    $mail->SMTPAuth = true;
    $mail->Username = "seu-email@gmail.com";  // seu e-mail
    $mail->Password = "sua-senha-de-aplicativo"; // senha de app do Gmail
    $mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;
    $mail->Port = 587;

    $mail->setFrom("seu-email@gmail.com", "Sistema de Relatórios");
    $mail->addAddress("destinatario@dominio.com");

    $mail->Subject = "Relatório Rodante (Excel)";
    $mail->Body = "Segue em anexo o relatório Excel.";

    $mail->addAttachment($arquivo);

    $mail->send();
    unlink($arquivo);
    echo "E-mail enviado com sucesso!";
} catch (Exception $e) {
    echo "Erro ao enviar e-mail: " . $mail->ErrorInfo;
}
