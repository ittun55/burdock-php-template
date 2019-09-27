<?php
namespace Burdock\Template\Renderer;

interface IRenderer
{
    static function render(string $template_path, array $config, array $data, string $output_path): void;
}