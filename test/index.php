<?php
    if (session_status() === PHP_SESSION_NONE) {
        session_start();
    }

    $flashNotification = null;
    if (!empty($_SESSION['flash_notification']) && is_array($_SESSION['flash_notification'])) {
        $flashNotification = $_SESSION['flash_notification'];
        unset($_SESSION['flash_notification']);
    }

    $flashData = is_array($flashNotification) ? $flashNotification : [];
    $flashDataType = (string) ($flashData['type'] ?? 'info');
    $flashType = in_array($flashDataType, ['success', 'error', 'info'], true) ? $flashDataType : 'info';
    $flashTitle = htmlspecialchars((string) ($flashData['title'] ?? ''), ENT_QUOTES, 'UTF-8');
    $flashMessage = htmlspecialchars((string) ($flashData['message'] ?? ''), ENT_QUOTES, 'UTF-8');
    $hasFlash = $flashNotification !== null;
    $flashMap = [
        'success' => ['bg' => '#ecfdf5', 'border' => '#10b981', 'title' => '#065f46'],
        'error' => ['bg' => '#fef2f2', 'border' => '#ef4444', 'title' => '#991b1b'],
        'info' => ['bg' => '#eff6ff', 'border' => '#3b82f6', 'title' => '#1d4ed8'],
    ];
    $flashStyle = $flashMap[$flashType ?? 'info'] ?? $flashMap['info'];
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Upload Course Material</title>
    <style>
        body {
            margin: 0;
            min-height: 100vh;
            font-family: Arial, Helvetica, sans-serif;
            background: linear-gradient(135deg, #f8fafc 0%, #eef2ff 100%);
            color: #0f172a;
        }

        .page {
            max-width: 920px;
            margin: 0 auto;
            padding: 48px 20px;
        }

        .card {
            background: rgba(255, 255, 255, 0.92);
            border: 1px solid rgba(148, 163, 184, 0.2);
            border-radius: 24px;
            box-shadow: 0 24px 60px rgba(15, 23, 42, 0.10);
            padding: 30px;
            backdrop-filter: blur(8px);
        }

        .title {
            margin: 0 0 10px;
            font-size: 34px;
            line-height: 1.1;
            letter-spacing: .02em;
        }

        .subtitle {
            margin: 0 0 24px;
            color: #475569;
            line-height: 1.6;
        }

        .upload-form {
            display: grid;
            gap: 14px;
            max-width: 560px;
        }

        .upload-form input[type="file"] {
            padding: 12px;
            border-radius: 12px;
            border: 1px solid #cbd5e1;
            background: #fff;
        }

        .upload-form button {
            width: fit-content;
            border: 0;
            border-radius: 999px;
            padding: 12px 20px;
            background: linear-gradient(135deg, #1d4ed8 0%, #4338ca 100%);
            color: #fff;
            font-weight: 700;
            cursor: pointer;
            box-shadow: 0 12px 24px rgba(37, 99, 235, 0.24);
        }

        .toast {
            position: fixed;
            right: 20px;
            bottom: 20px;
            z-index: 9999;
            width: min(380px, calc(100vw - 40px));
            padding: 16px 18px 14px;
            border-radius: 16px;
            border: 1px solid <?= $flashStyle['border'] ?>;
            background: <?= $flashStyle['bg'] ?>;
            box-shadow: 0 18px 45px rgba(15, 23, 42, 0.18);
            transform: translateY(20px);
            opacity: 0;
            transition: all .28s ease;
        }

        .toast.show {
            transform: translateY(0);
            opacity: 1;
        }

        .toast-title {
            margin: 0 24px 6px 0;
            font-size: 14px;
            font-weight: 700;
            color: <?= $flashStyle['title'] ?>;
        }

        .toast-message {
            margin: 0;
            font-size: 14px;
            line-height: 1.55;
            color: #334155;
        }

        .toast-close {
            position: absolute;
            top: 10px;
            right: 12px;
            border: 0;
            background: transparent;
            font-size: 18px;
            line-height: 1;
            cursor: pointer;
            color: #64748b;
        }

        @media (max-width: 640px) {
            .page {
                padding: 20px 14px;
            }

            .card {
                padding: 22px;
                border-radius: 20px;
            }

            .title {
                font-size: 28px;
            }
        }
    </style>
</head>
<body>
    <div class="page">
        <div class="card">
            <h1 class="title">Upload Course Material (PDF / Word / PPTX)</h1>
            <!-- <p class="subtitle">Upload a PDF, Word document, or PowerPoint file. The system will process it in the backend and show the result here as a popup.</p> -->

            <form class="upload-form" action="test.php" method="POST" enctype="multipart/form-data">
                <!-- <label for="pdf_file">Upload Course Material (PDF / Word / PPTX)</label> -->

                <input
                    type="file"
                    id="pdf_file"
                    name="pdf_file"
                    accept=".pdf,.doc,.docx,.pptx,application/pdf,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    required
                >

                <button type="submit">Upload</button>
            </form>
        </div>
    </div>

    <?php if ($hasFlash): ?>
        <div id="toast" class="toast" role="status" aria-live="polite">
            <button class="toast-close" type="button" aria-label="Close notification" onclick="hideToast()">&times;</button>
            <div class="toast-title"><?= $flashTitle ?></div>
            <div class="toast-message"><?= $flashMessage ?></div>
        </div>
        <script>
            const toast = document.getElementById('toast');

            function hideToast() {
                if (toast) {
                    toast.classList.remove('show');
                }
            }

            window.addEventListener('load', () => {
                if (!toast) {
                    return;
                }

                setTimeout(() => toast.classList.add('show'), 50);
                setTimeout(hideToast, 4500);
            });
        </script>
    <?php endif; ?>
</body>
</html>