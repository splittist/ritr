import type { FormattingTarget } from './direct-edit';
import { highlightColors } from '../engine/format';
export function FormattingToolbar({ target, busy }: { target?: FormattingTarget; busy: boolean }) {
  const disabled = busy || !target?.editable;
  return (
    <div className="format-toolbar" role="toolbar" aria-label="Text formatting">
      {(
        [
          ['b', 'Bold', 'B'],
          ['i', 'Italic', 'I'],
          ['u', 'Underline', 'U'],
        ] as const
      ).map(([key, label, text]) => (
        <button
          key={key}
          type="button"
          aria-label={label}
          aria-pressed={!!target?.values[key]}
          disabled={disabled}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => target?.apply(key)}
        >
          {text}
        </button>
      ))}
      <label>
        Color{' '}
        <input
          type="color"
          aria-label="Text color"
          disabled={disabled}
          value={
            /^[0-9a-f]{6}$/i.test(target?.values.color ?? '')
              ? `#${target!.values.color}`
              : '#000000'
          }
          onChange={(e) => target?.apply({ color: e.target.value.slice(1).toUpperCase() })}
        />
      </label>
      <button
        type="button"
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => target?.apply({ color: 'auto' })}
      >
        Automatic color
      </button>
      <label>
        Highlight{' '}
        <select
          aria-label="Text highlight"
          disabled={disabled}
          value={
            Object.hasOwn(highlightColors, target?.values.highlight ?? '')
              ? target!.values.highlight
              : 'none'
          }
          onChange={(e) => target?.apply({ highlight: e.target.value })}
        >
          <option value="none">None</option>
          {Object.keys(highlightColors).map((color) => (
            <option key={color} value={color}>
              {color.replace(/([A-Z])/g, ' $1')}
            </option>
          ))}
        </select>
      </label>
    </div>
  );
}
