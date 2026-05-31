/**
 * Excel Custom Ribbon Creator - Extended Emojis and Icon Registry (Korean Windows Style)
 * Classifies emojis into 5 distinct groups matching Windows (<Win + .>) keyboard structure.
 */

// --- 1. Windows Style Emoji Groups ---
const EMOJI_GROUPS = {
    "smileys": {
        name: "웃는 얼굴 및 감정",
        items: [
            "😀", "😃", "😄", "😁", "😆", "😅", "😂", "🤣", "😊", "😇", "🙂", "🙃", "😉", "😌", "😍", "🥰", "😘", 
            "😗", "😙", "😚", "😋", "😛", "😝", "😜", "🤪", "🤨", "🧐", "🤓", "😎", "🥸", "🤩", "🥳", "😏", 
            "😒", "😞", "😔", "😟", "😕", "🙁", "☹️", "😣", "😖", "😫", "😩", "🥺", "😢", "😭", "😤", "😠", 
            "😡", "🤬", "🤯", "😳", "🥵", "🥶", "😱", "😨", "😰", "😥", "😓", "🤗", "🤔", "🫣", "🤭", "🤫", 
            "🤥", "😶", "😐", "😑", "😬", "🫨", "🫠", "🫥", "😯", "😦", "😧", "😮", "😲", "🥱", "😴", "🤤", 
            "😪", "😵", "😵‍💫", "🫵", "👍", "👎", "👊", "✊", "🤛", "🤜", "🤞", "✌️", "🤟", "🤘", "👌", "🤌", 
            "🤏", "👈", "👉", "👆", "👇", "☝️", "✋", "🤚", "🖐️", "🖖", "👋", "🤙", "👏", "🙌", "🫶", "👐", 
            "🤲", "🤝", "🙏", "✍️", "💪"
        ]
    },
    "nature": {
        name: "동물 및 자연",
        items: [
            "🐶", "🐱", "🐭", "🐹", "🐰", "🦊", "🐻", "🐼", "🐨", "🐯", "🦁", "🐮", "🐷", "🐽", "🐸", "🐵", 
            "🙈", "🙉", "🙊", "🐒", "🐔", "🐧", "🐦", "🐤", "🐣", "🐥", "🦆", "🦅", "🦉", "🦇", "🐺", "🐗", 
            "🐴", "🦄", "🐝", "🐛", "🦋", "🐌", "🐞", "🐜", "🕷️", "🕸️", "🦂", "🐢", "🐍", "🦎", "🐙", "🦑", 
            " Lobster", "🦪", "🦀", "🐡", "🐠", "🐟", "🐬", "🐳", "🐋", "🦈", "🐊", "🐅", "🐆", "🦓", "🦍", "🦧", 
            "🐘", "🦛", "🦏", "🐪", "🐫", "🦒", "🦘", "🐂", "🐄", "🐎", "🐖", "🐏", "🐑", "🐐", "🦌", "🐕", 
            "🐈", "🐓", "🦃", "🦚", " Parrot", "🦢", "🦩", "🕊️", "🐇", "🐿️", "🦔", "🐾", "🐉", "🐲", "🌵", "🎄", 
            "🌲", "🌳", "🌴", "🌱", "🌿", "🍀", "🍁", "🍂", "🍃", "🍄", "🐚", "🌾", "💐", "🌷", "🌹", "🥀", 
            "🌺", "🌸", "🌼", "🌻", "🌞", "🌝", "🌙", "🌎", "🪐", "💫", "⭐", "🌟", "✨", "⚡", "☄️", "💥", 
            "🔥", "🌈", "☀️", "🌤", "🌥", "☁️", "🌧", "🌨", "🌩", "🌪", "🌊", "❄️", "🏔️", "🌋", "🗻"
        ]
    },
    "people": {
        name: "피플 및 활동",
        items: [
            "👤", "👥", "🫂", "🏃", "🚶", "💃", "🕺", "🕴️", "🧗", "🧘", "🏌️", "🏄", "🏊", "🚴", "🏇", "⛷️", 
            "🏂", "🏋️", "🤼", "🤸", "⛹️", "🤺", "🤾", "🏌️", "🏇", "🧘", "🧗", "🧗‍♂️", "🏆", "🥇", "🥈", "🥉", "🏅", 
            "🎖️", "🎫", "🎟️", "🎭", "🎨", "🎬", "🎤", "🎧", "🎼", "🎹", "🥁", "🎷", "🎺", "🎸", "🎻", "🎲", 
            "🎯", " bowling", "🎮", "🎰", "🧩", "👔", "🎓", "👑", "🎒", "💼", "👠", "👟", "🎩"
        ]
    },
    "objects": {
        name: "사물 및 물건",
        items: [
            "💻", "🖥️", "📱", "🔋", "🔌", "⚙️", "🛠️", "🔧", "🔨", "🔒", "🔓", "🔑", "📦", "✉️", "📧", "📝", 
            "📁", "📂", "📊", "📈", "📉", "💾", "🗑️", "🔔", "📎", "💼", "🧮", "⏰", "🔍", "💡", "🛡️", "💎", 
            "🏆", "🎯", "📓", "📕", "📔", "📖", "🔖", "🏷️", "🪙", "💵", "💳", "📯", "📥", "📤", "📮", "📦"
        ]
    },
    "symbols": {
        name: "기호 및 심볼",
        items: [
            "⚠️", "🚨", "🛑", "✅", "❌", "❓", "ℹ️", "➕", "➖", "✖️", "➗", "🟰", "🌐", "🔗", "♻️", "💬", 
            "📢", "📣", "🎵", "🎶", "💤", "❇️", "✳️", "✴️", "📶", "🔀", "🔁", "🔂", "🔄", "🔃", "◀️", "▶️", 
            "🔼", "🔽", "⏫", "⏬", "⏮️", "⏭️", "⏪", "⏩", "⏺️", "⏹️", "⏸️", "🎦", "📶", "📳", "📴", "🚹", 
            "🚺", "♿️", "🚰", "🚭", "🚫", "⛔️", "🔞", "☢️", "☣️", "⚜️", "🔱", "💯", "🔤", "🔢", "🔣"
        ]
    }
};

/**
 * Global fallback and emoji lookup renderer
 */
function getIconSvg(iconKey, size = 18) {
    if (!iconKey) return "";
    
    // Check if it's a raw Emoji
    const isEmoji = iconKey.length <= 4 && /\p{Emoji}/u.test(iconKey);
    if (isEmoji) {
        return `<span class="emoji-render-view" style="font-size: ${size}px; line-height:1; display:inline-block; font-family:'Segoe UI Emoji','Apple Color Emoji',sans-serif;">${iconKey}</span>`;
    }

    // Default basic Mso SVGs (embedded fallbacks if local files fail to load)
    const SVG_FALLBACKS = {
        "Folder": `<path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"></path>`,
        "Save": `<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"></path><polyline points="17 21 17 13 7 13 7 21" fill="none" stroke="currentColor" stroke-width="2"></polyline>`,
        "ChartInsert": `<line x1="18" y1="20" x2="18" y2="10" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"></line><line x1="12" y1="20" x2="12" y2="4" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"></line><line x1="6" y1="20" x2="6" y2="14" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"></line>`,
        "MacroPlay": `<polygon points="5 3 19 12 5 21 5 3" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"></polygon>`,
        "Settings": `<circle cx="12" cy="12" r="3" fill="none" stroke="currentColor" stroke-width="2"></circle><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"></path>`
    };

    const markup = SVG_FALLBACKS[iconKey] || `<circle cx="12" cy="12" r="10" fill="none" stroke="currentColor" stroke-width="2"></circle><text x="12" y="15" font-family="sans-serif" font-size="8" text-anchor="middle" fill="currentColor" font-weight="bold">M365</text>`;
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 24 24" class="mso-svg-icon">${markup}</svg>`;
}

// Export modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { EMOJI_GROUPS, getIconSvg };
} else {
    window.EMOJI_GROUPS = EMOJI_GROUPS;
    window.getIconSvg = getIconSvg;
}
