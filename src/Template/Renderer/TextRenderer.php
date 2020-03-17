<?php
namespace Burdock\Template\Renderer;

use Twig\Environment;
use Twig\Loader\ArrayLoader;
use Twig\Loader\FilesystemLoader;

class TextRenderer implements IRenderer
{
    static function render(string $template_path, array $config, array $data, string $output_path): void
    {
        $loader = new FilesystemLoader(dirname($template_path));
        $twig = new Environment($loader);
        $template = $twig->load(basename($template_path));
        $result = $template->render($data);
        file_put_contents($output_path, $result);
    }

    static function renderText(string $template, array $data): string
    {
        $loader = new ArrayLoader(['target' => $template]);
        $twig = new Environment($loader);
        return $twig->render('target', $data);
    }
}