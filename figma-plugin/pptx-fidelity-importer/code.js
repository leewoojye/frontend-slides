figma.showUI(__html__, { width: 460, height: 330, themeColors: true });

function hexToRgb(hex) {
  const clean = String(hex || '#202124').replace('#', '');
  const value = parseInt(clean.length === 3
    ? clean.split('').map((c) => c + c).join('')
    : clean, 16);
  return {
    r: ((value >> 16) & 255) / 255,
    g: ((value >> 8) & 255) / 255,
    b: (value & 255) / 255,
  };
}

function styleCandidates(layer) {
  const family = (layer.fontFamily || '').trim();
  const bold = Number(layer.fontWeight || 400) >= 600;
  const italic = Boolean(layer.italic);
  const requestedStyle = bold && italic ? 'Bold Italic' : bold ? 'Bold' : italic ? 'Italic' : 'Regular';
  const candidates = [];
  if (family) candidates.push({ family, style: requestedStyle });
  if (family && requestedStyle !== 'Regular') candidates.push({ family, style: 'Regular' });
  candidates.push({ family: 'Inter', style: requestedStyle });
  candidates.push({ family: 'Inter', style: bold ? 'Bold' : 'Regular' });
  candidates.push({ family: 'Inter', style: 'Regular' });
  return candidates;
}

async function loadBestFont(layer) {
  let lastError = null;
  for (const fontName of styleCandidates(layer)) {
    try {
      await figma.loadFontAsync(fontName);
      return fontName;
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError || new Error('No usable font found');
}

function horizontalAlign(value) {
  const v = String(value || '').toLowerCase();
  if (v === 'center') return 'CENTER';
  if (v === 'right') return 'RIGHT';
  if (v === 'justify') return 'JUSTIFIED';
  return 'LEFT';
}

function verticalAlign(value) {
  const v = String(value || '').toLowerCase();
  if (v === 'center') return 'CENTER';
  if (v === 'bottom') return 'BOTTOM';
  return 'TOP';
}

function appendAt(parent, node, x, y) {
  parent.appendChild(node);
  node.x = Number(x || 0);
  node.y = Number(y || 0);
  return node;
}

function addPlaceholderBox(slide, layer) {
  const rect = figma.createRectangle();
  appendAt(slide, rect, layer.x, layer.y);
  rect.resize(Math.max(1, Number(layer.width || 1)), Math.max(1, Number(layer.height || 1)));
  rect.name = `Placeholder box · ${layer.placeholderType || 'text'}`;
  rect.fills = [];
  rect.strokes = [{ type: 'SOLID', color: { r: 0.42, g: 0.42, b: 0.42 } }];
  rect.strokeWeight = 2;
  rect.dashPattern = [12, 8];
  rect.opacity = 0.34;
  rect.locked = false;
  return rect;
}

async function addTextLayer(slide, layer) {
  if (layer.placeholder) addPlaceholderBox(slide, layer);

  const text = figma.createText();
  appendAt(slide, text, layer.x, layer.y);
  text.name = layer.placeholder
    ? `Placeholder · ${layer.placeholderType || 'text'}`
    : layer.name || 'PPTX text';

  const fontName = await loadBestFont(layer);
  text.fontName = fontName;
  text.fontSize = Math.max(1, Number(layer.fontSize || 28));
  text.characters = String(layer.text || '');
  text.textAutoResize = 'NONE';
  text.resize(Math.max(1, Number(layer.width || 1)), Math.max(1, Number(layer.height || 1)));
  text.textAlignHorizontal = horizontalAlign(layer.align);
  text.textAlignVertical = verticalAlign(layer.verticalAlign);
  text.fills = [{
    type: 'SOLID',
    color: hexToRgb(layer.color),
    opacity: Math.max(0, Math.min(1, Number(layer.opacity ?? 1))),
  }];
  text.rotation = Number(layer.rotation || 0);
  return text;
}

async function addVectorBackground(slide, svg, deckWidth, deckHeight) {
  const node = figma.createNodeFromSvg(String(svg || ''));
  appendAt(slide, node, 0, 0);
  node.name = 'PPTX vector background';
  const width = Math.max(1, Number(deckWidth || 1920));
  const height = Math.max(1, Number(deckHeight || 1080));
  if (node.width !== width || node.height !== height) {
    node.resize(width, height);
  }
  node.locked = true;
  return node;
}

async function importDeck(deck) {
  if (!deck || !Array.isArray(deck.slides) || deck.slides.length === 0) {
    throw new Error('No slides found in the embedded manifest.');
  }

  const width = Number(deck.width || 1920);
  const height = Number(deck.height || 1080);
  const createdSlides = [];
  let fontFallbacks = 0;

  for (let index = 0; index < deck.slides.length; index += 1) {
    const source = deck.slides[index];
    const slide = figma.createSlide();
    slide.name = source.name || `Slide ${index + 1}`;

    await addVectorBackground(slide, source.svg, width, height);

    for (const layer of source.textLayers || []) {
      const requested = (layer.fontFamily || '').trim();
      const textNode = await addTextLayer(slide, layer);
      if (requested && textNode.fontName.family !== requested) fontFallbacks += 1;
    }

    createdSlides.push(slide);
  }

  if (createdSlides.length) {
    figma.viewport.scrollAndZoomIntoView(createdSlides);
    figma.currentPage.selection = createdSlides;
  }

  return { count: createdSlides.length, fontFallbacks };
}

figma.ui.onmessage = async (message) => {
  if (message?.type === 'cancel') {
    figma.closePlugin();
    return;
  }

  if (message?.type !== 'import-deck') return;

  try {
    figma.ui.postMessage({ type: 'status', message: 'Creating slides…' });
    const result = await importDeck(message.deck);
    figma.ui.postMessage({
      type: 'done',
      message: `Imported ${result.count} slide(s).${result.fontFallbacks ? ` ${result.fontFallbacks} text layer(s) used a fallback font.` : ''}`,
    });
    figma.notify(`Imported ${result.count} PowerPoint slide(s)`);
  } catch (error) {
    const messageText = error instanceof Error ? error.message : String(error);
    figma.ui.postMessage({ type: 'error', message: messageText });
    figma.notify(`PPTX import failed: ${messageText}`, { error: true });
  }
};
