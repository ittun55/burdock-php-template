<?php

use Burdock\Template\Renderer\TextRenderer;
use PHPUnit\Framework\TestCase;

class TextRendererTest extends TestCase
{
    public function testRender()
    {
        $tpl = "abc {{ def }} ghi";
        $tpl_path = __DIR__.'/tmp/template.txt';
        $out_path = __DIR__.'/tmp/output.txt';
        file_put_contents($tpl_path, $tpl);
        TextRenderer::render($tpl_path, [], ['def'=>'DEF'], $out_path);
        $this->assertEquals("abc DEF ghi", file_get_contents($out_path));
    }

    public function testRenderText()
    {
        $tpl = "abc {{ def }} ghi";
        $out = TextRenderer::renderText($tpl, ['def'=>'DEF']);
        $this->assertEquals("abc DEF ghi", $out);
    }
}
