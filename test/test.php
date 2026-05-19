<?php
    if (session_status() === PHP_SESSION_NONE) {
        session_start();
    }

    require 'vendor/autoload.php'; // for PHPMailer and any PDF library

    use PHPMailer\PHPMailer\PHPMailer;
    use PHPMailer\PHPMailer\Exception as MailException;
    use PhpOffice\PhpPresentation\Reader\PowerPoint2007 as PptxReader;
    use PhpOffice\PhpPresentation\Reader\PowerPoint97 as PptReader;
    use PhpOffice\PhpPresentation\Shape\RichText as PptRichText;
    use PhpOffice\PhpPresentation\Shape\Media as PptMedia;
    use thiagoalessio\TesseractOCR\TesseractOCR;

    function setFlashNotification(string $type, string $title, string $message): void {
        $_SESSION['flash_notification'] = [
            'type' => $type,
            'title' => $title,
            'message' => $message,
        ];
    }

    function redirectToSelf(): void {
        if (!headers_sent()) {
            header('Location: index.php');
            exit;
        }
    }

    function respondWithFlash(string $type, string $title, string $message): void {
        setFlashNotification($type, $title, $message);
        redirectToSelf();
    }

    function isPresentationMimeType(string $mimeType): bool {
        return in_array($mimeType, [
            'application/vnd.ms-powerpoint',
            'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        ], true);
    }

    function classifyMediaByPath(array &$media, string $path): void {
        $path = trim($path);
        if ($path === '') {
            return;
        }

        $resolvedPath = parse_url($path, PHP_URL_PATH);
        if (!is_string($resolvedPath) || $resolvedPath === '') {
            $resolvedPath = $path;
        }

        $extension = strtolower(pathinfo($resolvedPath, PATHINFO_EXTENSION));
        if (in_array($extension, ['mp4', 'webm', 'ogv', 'wmv', 'mov'], true)) {
            addUniqueMediaItem($media['videos'], $path);
            return;
        }
        if (in_array($extension, ['mp3', 'wav', 'ogg', 'm4a', 'aac'], true)) {
            addUniqueMediaItem($media['audio'], $path);
            return;
        }
        if (in_array($extension, ['jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp', 'svg'], true)) {
            addUniqueMediaItem($media['images'], $path);
            return;
        }

        addUniqueMediaItem($media['images'], $path);
    }

    function loadPresentationDocument(string $filePath, string $mimeType): object {
        if ($mimeType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation') {
            $reader = new PptxReader();
            return $reader->load($filePath);
        }

        $reader = new PptReader();
        return $reader->load($filePath);
    }

    function addUniqueMediaItem(array &$bucket, string $value): void {
        $value = trim($value);
        if ($value === '' || in_array($value, $bucket, true)) {
            return;
        }
        $bucket[] = $value;
    }

    function ensureDirectoryExists(string $dirPath): void {
        if (!is_dir($dirPath)) {
            mkdir($dirPath, 0777, true);
        }
    }

    function removeDirectoryRecursively(string $dirPath): void {
        if (!is_dir($dirPath)) {
            return;
        }

        $items = new RecursiveIteratorIterator(
            new RecursiveDirectoryIterator($dirPath, FilesystemIterator::SKIP_DOTS),
            RecursiveIteratorIterator::CHILD_FIRST
        );

        foreach ($items as $item) {
            if ($item->isDir()) {
                rmdir($item->getPathname());
            } else {
                unlink($item->getPathname());
            }
        }

        rmdir($dirPath);
    }

    function isDirectoryEmpty(string $dirPath): bool {
        if (!is_dir($dirPath)) {
            return true;
        }

        $entries = scandir($dirPath);
        if ($entries === false) {
            return false;
        }

        return count(array_diff($entries, ['.', '..'])) === 0;
    }

    function sanitizePathSegment(string $value): string {
        $value = preg_replace('/[^A-Za-z0-9._-]+/', '_', $value);
        return trim((string) $value, '._-');
    }

    function isImagePath(string $path): bool {
        $extension = strtolower(pathinfo($path, PATHINFO_EXTENSION));
        return in_array($extension, ['jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp', 'svg'], true);
    }

    function createMediaAssetStore(): array {
        return [
            'inlineImages' => [],
            'attachments' => [],
            'heroImage' => null,
        ];
    }

    function normalizeMatchingText(string $text): string {
        $text = strtolower($text);
        $text = preg_replace('/[^a-z0-9]+/', ' ', $text);
        $text = preg_replace('/\s+/', ' ', $text);
        return trim((string) $text);
    }

    function extractMatchingKeywords(string $text): array {
        $text = normalizeMatchingText($text);
        if ($text === '') {
            return [];
        }

        $stopWords = array_fill_keys([
            'the', 'and', 'for', 'with', 'that', 'this', 'from', 'into', 'your', 'have', 'has', 'was', 'were', 'are', 'you', 'will', 'been', 'can', 'any', 'all', 'new', 'course', 'material', 'module', 'lms', 'open', 'now', 'about', 'part', 'lab', 'virtual', 'machine', 'workstation'
        ], true);

        $tokens = preg_split('/\s+/', $text, -1, PREG_SPLIT_NO_EMPTY);
        if (!is_array($tokens)) {
            return [];
        }

        $keywords = [];
        foreach ($tokens as $token) {
            $token = trim($token);
            if ($token === '' || strlen($token) < 4 || isset($stopWords[$token])) {
                continue;
            }

            $keywords[$token] = true;
        }

        return array_keys($keywords);
    }

    function getImageOcrText(string $filePath): string {
        if (!class_exists('thiagoalessio\\TesseractOCR\\TesseractOCR') || !is_file($filePath)) {
            return '';
        }

        try {
            $ocr = new TesseractOCR($filePath);
            $text = $ocr->run();
            return normalizeMatchingText($text);
        } catch (\Throwable $e) {
            return '';
        }
    }

    function scoreMediaAssetForContext(array $asset, string $contextText): int {
        if (!is_array($asset) || empty($asset['path']) || !is_file($asset['path'])) {
            return 0;
        }

        $keywords = extractMatchingKeywords($contextText);
        if (empty($keywords)) {
            return 1;
        }

        $haystacks = [
            normalizeMatchingText($asset['name'] ?? ''),
            normalizeMatchingText(basename((string) ($asset['path'] ?? ''))),
            getImageOcrText($asset['path']),
        ];

        $score = 0;
        foreach ($keywords as $keyword) {
            foreach ($haystacks as $haystack) {
                if ($haystack !== '' && strpos($haystack, $keyword) !== false) {
                    $score += 3;
                }
            }
        }

        if ($score === 0 && !empty($haystacks[2])) {
            foreach (['error', 'virtualbox', 'network', 'lab', 'machine', 'setup', 'install', 'security', 'linux', 'window'] as $hint) {
                if (strpos($haystacks[2], $hint) !== false) {
                    $score += 2;
                }
            }
        }

        $score += min(5, (int) round((int) ($asset['size'] ?? 0) / 250000));
        return $score;
    }

    function selectBestHeroImageAsset(array $mediaAssets, string $summaryText, string $moduleName, string $materialTitle, string $fileContent): ?array {
        if (!is_array($mediaAssets) || empty($mediaAssets['inlineImages'])) {
            return null;
        }

        $contextText = trim($summaryText . ' ' . $moduleName . ' ' . $materialTitle . ' ' . $fileContent);
        $bestAsset = null;
        $bestScore = 0;

        foreach ($mediaAssets['inlineImages'] as $asset) {
            $score = scoreMediaAssetForContext($asset, $contextText);
            if ($score > $bestScore) {
                $bestScore = $score;
                $bestAsset = $asset;
            }
        }

        if ($bestAsset === null) {
            return $mediaAssets['inlineImages'][0] ?? null;
        }

        return $bestAsset;
    }

    function registerExtractedMediaAsset(array &$mediaAssets, string $filePath, string $displayName, string $sourceType): void {
        $filePath = $filePath;
        $displayName = trim($displayName);
        if ($filePath === '' || !is_file($filePath)) {
            return;
        }

        $asset = [
            'path' => $filePath,
            'name' => $displayName !== '' ? $displayName : basename($filePath),
            'sourceType' => $sourceType,
            'size' => filesize($filePath) ?: 0,
        ];

        $mediaAssets['attachments'][] = $asset;

        if (isImagePath($filePath)) {
            $mediaAssets['inlineImages'][] = $asset;
            if ($mediaAssets['heroImage'] === null) {
                $mediaAssets['heroImage'] = $asset;
            }
        }
    }

    function extractEmbeddedOfficeMedia(string $filePath, string $mimeType, array &$media, array &$mediaAssets): void {
        if ($mimeType !== 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' &&
            $mimeType !== 'application/vnd.openxmlformats-officedocument.presentationml.presentation') {
            return;
        }

        if (!class_exists('ZipArchive')) {
            return;
        }

        $zip = new \ZipArchive();
        if ($zip->open($filePath) !== true) {
            return;
        }

        $baseName = sanitizePathSegment(pathinfo($filePath, PATHINFO_FILENAME));
        $extractDir = dirname($filePath) . DIRECTORY_SEPARATOR . 'extracted_media' . DIRECTORY_SEPARATOR . $baseName;
        ensureDirectoryExists($extractDir);

        $mediaPrefix = $mimeType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
            ? 'ppt/media/'
            : 'word/media/';

        for ($i = 0; $i < $zip->numFiles; $i++) {
            $entry = $zip->getNameIndex($i);
            if (!is_string($entry) || strpos($entry, $mediaPrefix) !== 0) {
                continue;
            }

            $entryContents = $zip->getFromIndex($i);
            if ($entryContents === false) {
                continue;
            }

            $entryName = basename($entry);
            $safeEntryName = sanitizePathSegment($entryName);
            if ($safeEntryName === '') {
                $safeEntryName = 'media_' . $i;
            }

            $targetPath = $extractDir . DIRECTORY_SEPARATOR . $safeEntryName;
            if (file_exists($targetPath)) {
                $targetPath = $extractDir . DIRECTORY_SEPARATOR . uniqid($safeEntryName . '_', true);
            }

            if (file_put_contents($targetPath, $entryContents) === false) {
                continue;
            }

            $extension = strtolower(pathinfo($targetPath, PATHINFO_EXTENSION));
            if (isImagePath($targetPath)) {
                $kind = 'images';
            } elseif (in_array($extension, ['mp4', 'webm', 'ogv', 'wmv', 'mov'], true)) {
                $kind = 'videos';
            } elseif (in_array($extension, ['mp3', 'wav', 'ogg', 'm4a', 'aac'], true)) {
                $kind = 'audio';
            } else {
                $kind = 'images';
            }

            addUniqueMediaItem($media[$kind], 'embedded:' . $entry);
            registerExtractedMediaAsset($mediaAssets, $targetPath, $entryName, $kind);
        }

        $zip->close();
    }

    function buildMediaPromptSummary(array $media, array $mediaAssets): string {
        $lines = [];
        foreach (['images', 'videos', 'iframes', 'links', 'audio'] as $key) {
            $items = $media[$key] ?? [];
            $lines[] = strtoupper($key) . ': ' . count($items);
            foreach (array_slice($items, 0, 5) as $item) {
                $lines[] = '- ' . $item;
            }
        }

        if (!empty($mediaAssets['attachments'])) {
            $lines[] = 'Extracted files:';
            foreach (array_slice($mediaAssets['attachments'], 0, 6) as $asset) {
                $lines[] = '- ' . $asset['name'] . ' (' . $asset['sourceType'] . ')';
            }
        }

        return implode("\n", $lines);
    }

    function decodeGeminiJson(string $text): ?array {
        $text = trim($text);
        if ($text === '') {
            return null;
        }

        $decoded = json_decode($text, true);
        if (is_array($decoded)) {
            return $decoded;
        }

        if (preg_match('/\{.*\}/s', $text, $matches)) {
            $decoded = json_decode($matches[0], true);
            if (is_array($decoded)) {
                return $decoded;
            }
        }

        return null;
    }

    function buildDocumentGeneralSummary(string $fileContent, int $maxLength = 320): string {
        $text = trim($fileContent);
        if ($text === '') {
            return 'New course material has been uploaded to the LMS. Open it now for the key concepts, step-by-step guidance, and important references.';
        }

        $text = preg_replace('/\s+/', ' ', $text);
        if (!is_string($text)) {
            $text = '';
        }

        if ($text === '') {
            return 'New course material has been uploaded to the LMS. Open it now for the key concepts, step-by-step guidance, and important references.';
        }

        $sentences = preg_split('/(?<=[.!?])\s+/', $text, -1, PREG_SPLIT_NO_EMPTY);
        if (!is_array($sentences) || empty($sentences)) {
            $sentences = [$text];
        }

        $selected = [];
        $length = 0;
        foreach ($sentences as $sentence) {
            $sentence = trim($sentence);
            if ($sentence === '') {
                continue;
            }

            $nextLength = $length + strlen($sentence) + (empty($selected) ? 0 : 1);
            if ($nextLength > $maxLength) {
                break;
            }

            $selected[] = $sentence;
            $length = $nextLength;

            if (count($selected) >= 2) {
                break;
            }
        }

        $summary = trim(implode(' ', $selected));
        if ($summary === '') {
            $summary = substr($text, 0, $maxLength);
        }

        if (strlen($summary) < strlen($text) && substr($summary, -1) !== '.') {
            $summary .= '...';
        }

        return $summary;
    }

    function detectModuleName(string $fileContent, string $filePath): string {
        $lookupText = strtolower(trim($fileContent) . ' ' . basename($filePath));
        $moduleMap = [
            'cyberops' => 'Module: CyberOps Virtual Lab Setup',
            'virtualbox' => 'Module: Virtualization and Lab Setup',
            'virtual machine' => 'Module: Virtual Machine Deployment',
            'vm' => 'Module: Virtual Machine Deployment',
            'network' => 'Module: Networking Foundations',
            'security' => 'Module: Cybersecurity Foundations',
            'linux' => 'Module: Linux Lab Essentials',
            'windows' => 'Module: Windows Lab Essentials',
            'html' => 'Module: Web Content and Publishing',
            'pdf' => 'Module: Document Processing Workflow',
            'word' => 'Module: Document Processing Workflow',
            'powerpoint' => 'Module: Presentation Workflow',
            'pptx' => 'Module: Presentation Workflow',
        ];

        foreach ($moduleMap as $keyword => $moduleName) {
            if (strpos($lookupText, $keyword) !== false) {
                return $moduleName;
            }
        }

        return 'Module: Course Material Update';
    }

    function extractMaterialTitle(string $fileContent, string $filePath): string {
        $text = trim($fileContent);
        if ($text !== '') {
            $lines = preg_split('/\R+/', $text, -1, PREG_SPLIT_NO_EMPTY);
            if (is_array($lines)) {
                foreach ($lines as $line) {
                    $line = trim($line);
                    if ($line === '') {
                        continue;
                    }

                    $line = preg_replace('/\s+/', ' ', $line);
                    if (strlen($line) > 18) {
                        return mb_substr($line, 0, 120);
                    }
                }
            }
        }

        $fallbackName = pathinfo((string) $filePath, PATHINFO_FILENAME);
        $fallbackName = preg_replace('/[_-]+/', ' ', (string) $fallbackName);
        $fallbackName = trim(preg_replace('/\s+/', ' ', $fallbackName));
        if ($fallbackName !== '') {
            return mb_substr($fallbackName, 0, 120);
        }

        return 'Uploaded Course Material';
    }

    function buildNotificationEmailBody(array $notification, string $contentPreview, string $mediaHtml, array $mediaAssets, ?string $heroCid = null, string $moduleName = '', string $materialTitle = ''): string {
        $heroSection = '';
        if ($heroCid !== null && !empty($mediaAssets['heroImage']['name'])) {
            $heroSection = '<div style="margin:24px 0 28px;"><img src="cid:' . htmlspecialchars($heroCid, ENT_QUOTES, 'UTF-8') . '" alt="' . htmlspecialchars($mediaAssets['heroImage']['name'], ENT_QUOTES, 'UTF-8') . '" style="width:100%;max-width:620px;border-radius:20px;display:block;border:0;box-shadow:0 18px 50px rgba(15,23,42,0.18);" /></div>';
        }

        $bridgeText = trim((string) $moduleName) !== '' && trim((string) $materialTitle) !== ''
            ? 'From ' . trim((string) $moduleName) . ' to ' . trim((string) $materialTitle) . ', this update is designed to move you forward step by step.'
            : 'This update is designed to move you forward step by step.';

        $bulletHtml = '';
        if (!empty($notification['bullets'])) {
            $bulletHtml = '<ul style="margin:18px 0 0;padding-left:20px;">';
            foreach ($notification['bullets'] as $bullet) {
                $bulletHtml .= '<li style="margin:0 0 10px;">' . htmlspecialchars($bullet, ENT_QUOTES, 'UTF-8') . '</li>';
            }
            $bulletHtml .= '</ul>';
        }

        return '<div style="margin:0;padding:0;background:linear-gradient(135deg,#dbeafe 0%,#e0e7ff 50%,#f8fafc 100%);font-family:Arial,Helvetica,sans-serif;">
            <div style="max-width:760px;margin:0 auto;padding:32px 18px;">
                <div style="background:#ffffff;border-radius:28px;padding:34px 30px;box-shadow:0 24px 60px rgba(15,23,42,0.12);border:1px solid rgba(148,163,184,0.18);overflow:hidden;">
                    <div style="display:inline-block;padding:7px 12px;border-radius:999px;background:#eef2ff;color:#3730a3;font-size:12px;letter-spacing:.08em;text-transform:uppercase;font-weight:700;">Just dropped in LMS</div>
                    <h1 style="margin:14px 0 10px;font-size:34px;line-height:1.15;letter-spacing:.06em;text-transform:uppercase;color:#6366f1;font-weight:700;">' . htmlspecialchars($notification['headline'], ENT_QUOTES, 'UTF-8') . '</h1>
                    <div style="margin:0 0 14px;font-size:15px;line-height:1.75;color:#0f172a;font-weight:600;">' . htmlspecialchars($bridgeText, ENT_QUOTES, 'UTF-8') . '</div>
                    <p style="font-size:17px;line-height:1.75;color:#334155;margin:0;">' . nl2br(htmlspecialchars($notification['summary'], ENT_QUOTES, 'UTF-8')) . '</p>'
                    . $heroSection
                    . $bulletHtml
                    . '<div style="margin-top:24px;padding:20px 22px;background:linear-gradient(135deg,#0f172a 0%,#1d4ed8 100%);color:#ffffff;border-radius:18px;">'
                    . '<strong style="display:block;font-size:20px;line-height:1.35;margin-bottom:8px;">' . htmlspecialchars($notification['cta'], ENT_QUOTES, 'UTF-8') . '</strong>'
                    . '</div>'
                . '</div>
            </div>
        </div>';
    }

    function safeCallMethod(mixed $object, string $methodName): mixed {
        if (!is_object($object) || !is_string($methodName) || !is_callable([$object, $methodName])) {
            return null;
        }

        return call_user_func([$object, $methodName]);
    }

    function extractMediaMetadata(string $filePath, string $mimeType, ?array &$mediaAssets = null): array {
        $media = [
            'images' => [],
            'videos' => [],
            'iframes' => [],
            'links' => [],
            'audio' => [],
        ];

        if (!is_array($mediaAssets)) {
            $mediaAssets = createMediaAssetStore();
        }

        if ($mimeType === 'text/html' || $mimeType === 'application/xhtml+xml') {
            $html = @file_get_contents($filePath);
            if ($html !== false) {
                $dom = new \DOMDocument();
                libxml_use_internal_errors(true);
                $dom->loadHTML($html);
                libxml_clear_errors();

                foreach ($dom->getElementsByTagName('img') as $node) {
                    addUniqueMediaItem($media['images'], $node->getAttribute('src'));
                }
                foreach ($dom->getElementsByTagName('video') as $node) {
                    addUniqueMediaItem($media['videos'], $node->getAttribute('src'));
                }
                foreach ($dom->getElementsByTagName('source') as $node) {
                    $src = $node->getAttribute('src');
                    if ($src !== '') {
                        $type = strtolower($node->getAttribute('type'));
                        if (strpos($type, 'audio') !== false) {
                            addUniqueMediaItem($media['audio'], $src);
                        } else {
                            addUniqueMediaItem($media['videos'], $src);
                        }
                    }
                }
                foreach ($dom->getElementsByTagName('iframe') as $node) {
                    addUniqueMediaItem($media['iframes'], $node->getAttribute('src'));
                }
                foreach ($dom->getElementsByTagName('a') as $node) {
                    addUniqueMediaItem($media['links'], $node->getAttribute('href'));
                }
                foreach ($dom->getElementsByTagName('audio') as $node) {
                    addUniqueMediaItem($media['audio'], $node->getAttribute('src'));
                }
            }
        }

        if (isPresentationMimeType($mimeType)) {
            try {
                $presentation = loadPresentationDocument($filePath, $mimeType);
                foreach ($presentation->getAllSlides() as $slide) {
                    foreach ($slide->getShapeCollection() as $shape) {
                        $shapePath = safeCallMethod($shape, 'getPath');
                        if ($shapePath !== null && ($shape instanceof PptMedia || is_string($shapePath))) {
                            classifyMediaByPath($media, (string) $shapePath);
                        }

                        $hyperlink = safeCallMethod($shape, 'getHyperlink');
                        if (is_object($hyperlink)) {
                            $url = safeCallMethod($hyperlink, 'getUrl');
                            if (is_string($url)) {
                                $url = trim($url);
                                if ($url !== '' && preg_match('/^https?:\/\//i', $url)) {
                                    addUniqueMediaItem($media['links'], $url);
                                }
                            }
                        }

                        if ($shape instanceof PptRichText) {
                            foreach ($shape->getParagraphs() as $paragraph) {
                                foreach ($paragraph->getRichTextElements() as $element) {
                                    $textHyperlink = safeCallMethod($element, 'getHyperlink');
                                    if (is_object($textHyperlink)) {
                                        $textUrl = safeCallMethod($textHyperlink, 'getUrl');
                                        if (is_string($textUrl)) {
                                            $textUrl = trim($textUrl);
                                            if ($textUrl !== '' && preg_match('/^https?:\/\//i', $textUrl)) {
                                                addUniqueMediaItem($media['links'], $textUrl);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            } catch (\Throwable $e) {
                // Keep upload flow alive even if some presentation metadata cannot be parsed.
            }

            extractEmbeddedOfficeMedia($filePath, $mimeType, $media, $mediaAssets);
        }

        if ($mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' && class_exists('ZipArchive')) {
            $zip = new \ZipArchive();
            if ($zip->open($filePath) === true) {
                $relsXml = $zip->getFromName('word/_rels/document.xml.rels');
                if ($relsXml !== false) {
                    $rels = @simplexml_load_string($relsXml);
                    if ($rels !== false) {
                        foreach ($rels->Relationship as $relationship) {
                            $type = (string) $relationship['Type'];
                            $target = (string) $relationship['Target'];
                            $targetMode = (string) $relationship['TargetMode'];

                            if (strpos($type, 'hyperlink') !== false && strtolower($targetMode) === 'external') {
                                addUniqueMediaItem($media['links'], $target);
                            }
                        }
                    }
                }

                for ($i = 0; $i < $zip->numFiles; $i++) {
                    $entry = $zip->getNameIndex($i);
                    if (strpos($entry, 'word/media/') === 0) {
                        addUniqueMediaItem($media['images'], 'embedded:' . $entry);
                    }
                }
                $zip->close();
            }

            extractEmbeddedOfficeMedia($filePath, $mimeType, $media, $mediaAssets);
        }

        if ($mimeType === 'application/pdf') {
            $pdfRaw = @file_get_contents($filePath);
            if ($pdfRaw !== false && preg_match_all('/https?:\/\/[^\s\)\]\}>"\']+/i', $pdfRaw, $matches)) {
                foreach ($matches[0] as $url) {
                    addUniqueMediaItem($media['links'], $url);
                }
            }
        }

        return $media;
    }

    function buildMediaHtml(array $media): string {
        $labelMap = [
            'images' => 'Images',
            'videos' => 'Videos',
            'iframes' => 'Iframes',
            'links' => 'Links',
            'audio' => 'Audio',
        ];

        $html = '<h3>Detected media and links</h3>';
        foreach ($labelMap as $key => $label) {
            $items = $media[$key] ?? [];
            $html .= '<p><strong>' . htmlspecialchars($label) . ' (' . count($items) . ')</strong></p>';

            if (empty($items)) {
                $html .= '<p>None found.</p>';
                continue;
            }

            $html .= '<ul>';
            foreach ($items as $item) {
                $safeText = htmlspecialchars($item);
                if (strpos($item, 'http://') === 0 || strpos($item, 'https://') === 0) {
                    $safeHref = htmlspecialchars($item);
                    $html .= '<li><a href="' . $safeHref . '" target="_blank" rel="noopener noreferrer">' . $safeText . '</a></li>';
                } else {
                    $html .= '<li>' . $safeText . '</li>';
                }
            }
            $html .= '</ul>';
        }

        return $html;
    }

    // 1. Function to read PDF or Word content
    function extractFileContent(string $filePath, string $mimeType) {
        if ($mimeType === 'application/pdf') {
            // PDF
            // Instantiate parser dynamically to avoid static analyzer undefined-type errors
            $parserClass = 'Smalot\\PdfParser\\Parser';
            if (!class_exists($parserClass)) {
                throw new \Exception("PDF parser library is not installed.");
            }
            $parser = new $parserClass();
            $pdf = $parser->parseFile($filePath);
            $text = $pdf->getText();

            if (empty(trim($text))) {
                throw new \Exception("Failed to extract text from PDF.");
            }
            return $text;

        } elseif (
            $mimeType === 'application/msword' ||
            $mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        ) {
            // Word (doc or docx)
            $ioFactoryClass = 'PhpOffice\\PhpWord\\IOFactory';
            if (!class_exists($ioFactoryClass)) {
                throw new \Exception("PhpWord library is not installed.");
            }
            $phpWord = $ioFactoryClass::load($filePath);
            $text = '';

            foreach ($phpWord->getSections() as $section) {
                foreach ($section->getElements() as $element) {
                    $textGetter = [$element, 'getText'];
                    if (is_callable($textGetter)) {
                        $text .= (string) call_user_func($textGetter) . "\n";
                    }
                }
            }

            if (empty(trim($text))) {
                throw new \Exception("Failed to extract text from Word file.");
            }
            return $text;

        } elseif ($mimeType === 'text/html' || $mimeType === 'application/xhtml+xml') {
            $html = @file_get_contents($filePath);
            if ($html === false) {
                throw new \Exception("Failed to read HTML file.");
            }

            $text = trim(html_entity_decode(strip_tags($html), ENT_QUOTES | ENT_HTML5, 'UTF-8'));
            if ($text === '') {
                throw new \Exception("Failed to extract text from HTML file.");
            }
            return $text;

        } elseif (isPresentationMimeType($mimeType)) {
            try {
                $presentation = loadPresentationDocument($filePath, $mimeType);
            } catch (\Throwable $e) {
                throw new \Exception("Failed to read PowerPoint file.");
            }

            $text = '';
            foreach ($presentation->getAllSlides() as $slide) {
                foreach ($slide->getShapeCollection() as $shape) {
                    if ($shape instanceof PptRichText) {
                        $text .= $shape->getPlainText() . "\n";
                        continue;
                    }

                    $textGetter = [$shape, 'getText'];
                    if (is_callable($textGetter)) {
                        $text .= (string) call_user_func($textGetter) . "\n";
                    }
                }
            }

            if (empty(trim($text))) {
                throw new \Exception("Failed to extract text from PowerPoint file.");
            }
            return $text;

        } else {
            throw new \Exception("Unsupported file type for extraction.");
        }
    }

    // 2. Function to generate a short notification using Gemini LLM
    function generateNotification(string $fileContent, array $media, array $mediaAssets, string $moduleName, string $materialTitle): array {
        
        $apiKey = 'YOUR_GEMINI_API_KEY'; // Replace with your Gemini API key

        $endpoint = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' . urlencode($apiKey);

        $mediaSummary = buildMediaPromptSummary($media, $mediaAssets);
        $contentExcerpt = substr($fileContent, 0, 12000);
        $bridgeSummary = trim((string) $moduleName) !== '' && trim((string) $materialTitle) !== ''
            ? trim((string) $moduleName) . ' and ' . trim((string) $materialTitle)
            : 'the course material';

        $prompt = "You are writing a connected LMS notification email with a clear learning flow.\n"
            . "Goal: make students feel curious, confident, and eager to open the LMS immediately.\n"
            . "Return strict JSON with these keys only: subject, preheader, headline, summary, bullets, cta, closing.\n"
            . "Writing rules:\n"
            . "1) Tone: energetic, smooth, student-friendly, and connected.\n"
            . "2) Subject: <= 70 chars, action oriented, and relevant to the module/material.\n"
            . "3) Preheader: <= 80 chars, build anticipation and flow into the main idea.\n"
            . "4) Headline: 5-10 words, bold and motivating.\n"
            . "5) Summary: 2-4 short sentences in one smooth flow. Start by connecting the module and material title, then explain what students will learn, and end with a natural reason to open LMS now.\n"
            . "6) Bullets: exactly 3 concise bullets; each bullet should feel like the next step in a learning journey.\n"
            . "7) CTA: one direct command that makes students want to continue immediately.\n"
            . "8) Closing: one short encouraging line that keeps the momentum.\n"
            . "9) Use the module and material names naturally in the copy.\n"
            . "10) Only mention media if it is supported by evidence below. Do not invent media.\n"
            . "11) Never reference this prompt or output formatting in the final text.\n\n"
            . "Module name:\n" . $moduleName . "\n\n"
            . "Material title:\n" . $materialTitle . "\n\n"
            . "Connected flow focus:\n" . $bridgeSummary . "\n\n"
            . "Media evidence:\n" . $mediaSummary . "\n\n"
            . "Document excerpt:\n" . $contentExcerpt;

        $postData = [
            'contents' => [
                [
                    'parts' => [
                        ['text' => $prompt]
                    ]
                ]
            ],
            'generationConfig' => [
                'maxOutputTokens' => 360,
                'temperature' => 0.9
            ]
        ];

        $requestBody = json_encode($postData);
        if ($requestBody === false) {
            throw new \Exception('Failed to encode Gemini request payload.');
        }

        if (function_exists('curl_init')) {
            $ch = curl_init($endpoint);
            curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
            curl_setopt($ch, CURLOPT_HTTPHEADER, [
                'Content-Type: application/json'
            ]);
            curl_setopt($ch, CURLOPT_POST, true);
            curl_setopt($ch, CURLOPT_POSTFIELDS, $requestBody);

            $response = curl_exec($ch);
            if ($response === false || curl_errno($ch)) {
                throw new \Exception('Gemini API Request Error: ' . curl_error($ch));
            }
            // curl_close is deprecated on modern PHP; releasing the handle is sufficient.
            $ch = null;
        } else {
            $context = stream_context_create([
                'http' => [
                    'method' => 'POST',
                    'header' => "Content-Type: application/json\r\n",
                    'content' => $requestBody,
                    'ignore_errors' => true,
                    'timeout' => 30,
                ],
            ]);

            $response = @file_get_contents($endpoint, false, $context);
            if ($response === false) {
                throw new \Exception('Gemini API Request Error: HTTP request failed without cURL extension.');
            }
        }

        $data = json_decode($response, true);
        if (isset($data['error']['message'])) {
            throw new \Exception('Gemini API Error: ' . $data['error']['message']);
        }

        $generatedText = $data['candidates'][0]['content']['parts'][0]['text'] ?? '';
        $generatedData = decodeGeminiJson($generatedText);

        if (!is_array($generatedData)) {
            return [
                'subject' => 'New Course Material Available for ' . (trim((string) $moduleName) !== '' ? trim((string) $moduleName) : 'Your Course'),
                'preheader' => 'A fresh course update just landed. Dive in now.',
                'headline' => trim((string) $moduleName) !== '' ? trim((string) $moduleName) : 'New material just dropped for your course',
                'summary' => trim($generatedText) !== '' ? trim($generatedText) : buildDocumentGeneralSummary($fileContent),
                'bullets' => [
                    'Get step-by-step guidance for faster completion of this topic.',
                    'Strengthen practical skills you can apply in upcoming tasks.',
                    'Reduce confusion by following a clear, structured learning flow.',
                ],
                'cta' => 'Open the LMS now and start this material today.',
                'closing' => 'You are one focused session away from a big improvement.',
            ];
        }

        $normalizedBullets = array_values(array_filter((array) ($generatedData['bullets'] ?? []), 'is_string'));
        if (count($normalizedBullets) > 3) {
            $normalizedBullets = array_slice($normalizedBullets, 0, 3);
        }

        if (count($normalizedBullets) < 3) {
            $fallbackBullets = [
                'Understand the core idea quickly with a clear learning path.',
                'Practice with concrete steps you can apply right away.',
                'Build confidence before your next class activity or assessment.',
            ];
            $normalizedBullets = array_slice(array_merge($normalizedBullets, $fallbackBullets), 0, 3);
        }

        return [
            'subject' => trim((string) ($generatedData['subject'] ?? ('New Course Material Available for ' . (trim((string) $moduleName) !== '' ? trim((string) $moduleName) : 'Your Course')))),
            'preheader' => trim((string) ($generatedData['preheader'] ?? 'A fresh course update just landed. Dive in now.')),
            'headline' => trim((string) ($generatedData['headline'] ?? (trim((string) $moduleName) !== '' ? $moduleName : 'New material just dropped for your course'))),
            'summary' => trim((string) ($generatedData['summary'] ?? buildDocumentGeneralSummary($fileContent))),
            'bullets' => $normalizedBullets,
            'cta' => trim((string) ($generatedData['cta'] ?? 'Open the LMS now and start this material today.')),
            'closing' => trim((string) ($generatedData['closing'] ?? 'You are one focused session away from a big improvement.')),
        ];
    }

    // 3. Function to send email to all enrolled students
    function sendEmailNotification(string $subject, string $message, array $recipients, array $mediaAssets = []): bool {
        $mail = new PHPMailer(true);

        try {
            // Server settings
            $mail->isSMTP();
            $mail->Host       = 'smtp.gmail.com'; // Replace with your SMTP server
            $mail->SMTPAuth   = true;
            $mail->Username   = 'YOUR_EMAIL@example.com';// Replace with your email
            $mail->Password   = 'YOUR_APP_PASSWORD'; // Replace with app password
            $mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;
            $mail->Port       = 587;
            $mail->CharSet    = 'UTF-8';

            $mail->setFrom('nithila0411@gmail.com', 'LMS Notification');

            $heroImagePath = null;
            if (is_array($mediaAssets) && !empty($mediaAssets['heroImage']['path']) && is_file($mediaAssets['heroImage']['path'])) {
                $heroImagePath = $mediaAssets['heroImage']['path'];
                $heroImageCid = 'hero-media-' . md5($heroImagePath) . '@lms';
                $mail->addEmbeddedImage($heroImagePath, $heroImageCid, basename($heroImagePath));
                $message = str_replace('cid:hero-media', 'cid:' . $heroImageCid, $message);
            }

            // Keep extracted media inline in the email body only; do not add downloadable attachments.

            foreach ($recipients as $email) {
                $mail->addAddress($email);
            }

            // Content
            $mail->isHTML(true);
            $mail->Subject = $subject;
            $mail->Body    = $message;

            $mail->send();
            return true;
        } catch (MailException $e) {
            error_log("Email could not be sent. Mailer Error: {$mail->ErrorInfo}");
            return false;
        }
    }

    // ===== Example usage after PDF upload =====

    // Assume $_FILES['pdf_file'] is uploaded
    if (isset($_FILES['pdf_file'])) {
        if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
            respondWithFlash('error', 'Upload failed', 'Invalid request.');
        }

        if (!isset($_FILES['pdf_file'])) {
            respondWithFlash('error', 'Upload failed', 'No file uploaded.');
        }

        if ($_FILES['pdf_file']['error'] !== UPLOAD_ERR_OK) {
            respondWithFlash('error', 'Upload failed', 'Upload error: ' . $_FILES['pdf_file']['error']);
        }

        $allowedMimeTypes = [
            'application/pdf',
            'application/msword',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'application/vnd.ms-powerpoint',
            'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            'text/html',
            'application/xhtml+xml'
        ];

        $finfo = new \finfo(FILEINFO_MIME_TYPE);
        if ($finfo === false) {
            respondWithFlash('error', 'Upload Failed', 'Could not Initialize File Info Detector.');
        }
        $mimeType = $finfo->file($_FILES['pdf_file']['tmp_name']);

        if (!in_array($mimeType, $allowedMimeTypes)) {
            respondWithFlash('error', 'Upload Failed', 'Invalid File Type.');
        }
        $uploadDir = 'uploads/';
        if (!is_dir($uploadDir)) {
            mkdir($uploadDir, 0777, true);
        }

        $fileName = time() . '_' . basename($_FILES['pdf_file']['name']);
        $filePath = $uploadDir . $fileName;

        if (move_uploaded_file($_FILES['pdf_file']['tmp_name'], $filePath)) {
            // Initialize defaults to avoid undefined variable warnings
            $fileText = '';
            $moduleName = '';
            $materialTitle = '';
            $mediaAssets = createMediaAssetStore();
            $mediaData = [];
            $extractedMediaDir = dirname($filePath) . DIRECTORY_SEPARATOR . 'extracted_media' . DIRECTORY_SEPARATOR . sanitizePathSegment(pathinfo($filePath, PATHINFO_FILENAME));

            try {
                // Extract text from PDF or Word
                $fileText = extractFileContent($filePath, $mimeType);
                $moduleName = detectModuleName($fileText, $filePath);
                $materialTitle = extractMaterialTitle($fileText, $filePath);
                $mediaAssets = createMediaAssetStore();
                $mediaData = extractMediaMetadata($filePath, $mimeType, $mediaAssets);
            } catch (\Exception $e) {
                error_log('Content extraction failed: ' . $e->getMessage());
                respondWithFlash('error', 'Upload failed', 'Failed to process uploaded content.');
            }

            // 2. Generate notification
            try {
                $notification = generateNotification($fileText, $mediaData, $mediaAssets, $moduleName, $materialTitle);
            } catch (\Exception $e) {
                $notification = [
                    'subject' => 'New Course Material Available for ' . (trim((string) $moduleName) !== '' ? trim((string) $moduleName) : 'Your Course'),
                    'preheader' => 'A new course material was shared in the LMS.',
                    'headline' => $moduleName,
                    'summary' => buildDocumentGeneralSummary($fileText),
                    'bullets' => [],
                    'cta' => 'Log into the LMS to view the latest material.',
                ];
                error_log('Notification generation failed: ' . $e->getMessage());
            }

            $selectedHeroImage = selectBestHeroImageAsset($mediaAssets, $notification['summary'] ?? '', $moduleName, $materialTitle, $fileText);
            if (is_array($selectedHeroImage)) {
                $mediaAssets['heroImage'] = $selectedHeroImage;
            }

            $contentPreview = substr($fileText, 0, 2000);
            $emailBody = buildNotificationEmailBody($notification, $contentPreview, '', $mediaAssets, 'hero-media', $moduleName, $materialTitle);

            // 3. Fetch enrolled students emails from your database
            $recipients = [
                'student1@example.com',
                'student2@example.com',
            ]; // Replace with dynamic query from LMS DB

            // 4. Send notification
            $subject = $notification['subject'] ?: "New Course Material Available!";
            if (sendEmailNotification($subject, $emailBody, $recipients, $mediaAssets)) {
                removeDirectoryRecursively($extractedMediaDir);
                $extractedMediaRoot = dirname($extractedMediaDir);
                if (isDirectoryEmpty($extractedMediaRoot)) {
                    rmdir($extractedMediaRoot);
                }
                respondWithFlash('success', 'Notification sent', 'Your notification was sent successfully.');
            } else {
                respondWithFlash('error', 'Notification failed', 'Failed to send notification.');
            }
        } else {
            respondWithFlash('error', 'Upload failed', 'Failed to upload the file.');
        }
    }

?>