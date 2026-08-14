/**
 * CellEditOverlay.tsx
 *
 * In-cell editing overlay for double-click / F2 / type-to-edit.
 * Supports Excel-like formula mode with cell reference picking.
 */

import React, { useEffect, useRef } from 'react';
import type { Address } from '@cyber-sheet/core';
import { isFormulaEditValue } from '../utils/formulaEdit';

export interface CellEditOverlayProps {
  cell: Address;
  /** Current edit buffer (controlled). */
  value: string;
  /** Viewport bounds from renderer.getCellBounds (already zoom-adjusted). */
  bounds: { x: number; y: number; width: number; height: number };
  cursorPosition?: number;
  onCommit: (value: string) => void;
  onCancel: () => void;
  onValueChange?: (value: string) => void;
  onCursorChange?: (position: number) => void;
  onNavigateVertical?: (direction: 'up' | 'down') => void;
  isPickingReference?: boolean;
  onReferencePickingChange?: (picking: boolean) => void;
  fontSize?: number;
}

export const CellEditOverlay: React.FC<CellEditOverlayProps> = ({
  cell,
  value,
  bounds,
  cursorPosition,
  onCommit,
  onCancel,
  onValueChange,
  onCursorChange,
  onNavigateVertical,
  isPickingReference = false,
  onReferencePickingChange,
  fontSize = 11,
}) => {
  const inputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    if (!inputRef.current) return;
    inputRef.current.focus();
    const pos = cursorPosition ?? value.length;
    inputRef.current.selectionStart = pos;
    inputRef.current.selectionEnd = pos;
  }, [cell.row, cell.col]);

  useEffect(() => {
    if (!inputRef.current || cursorPosition === undefined) return;
    if (
      inputRef.current.selectionStart === cursorPosition &&
      inputRef.current.selectionEnd === cursorPosition
    ) {
      return;
    }
    inputRef.current.selectionStart = cursorPosition;
    inputRef.current.selectionEnd = cursorPosition;
  }, [cursorPosition, value]);

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      e.stopPropagation();
      onCommit(value);
      return;
    }
    if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      onCancel();
      return;
    }
    if (e.key === 'Tab') {
      e.preventDefault();
      e.stopPropagation();
      onCommit(value);
      return;
    }
    if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
      e.preventDefault();
      e.stopPropagation();
      onNavigateVertical?.(e.key === 'ArrowUp' ? 'up' : 'down');
      return;
    }
    e.stopPropagation();
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const newValue = e.target.value;
    const cursor = e.target.selectionStart ?? newValue.length;
    onCursorChange?.(cursor);
    onValueChange?.(newValue);
    onReferencePickingChange?.(isFormulaEditValue(newValue));
  };

  const handleSelect = () => {
    const pos = inputRef.current?.selectionStart ?? value.length;
    onCursorChange?.(pos);
  };

  const handleBlur = (e: React.FocusEvent<HTMLInputElement>) => {
    const related = e.relatedTarget as HTMLElement | null;
    if (related?.closest('.formula-bar-input, .formula-bar')) return;
    if (isPickingReference) return;
    onCommit(value);
  };

  return (
    <input
      ref={inputRef}
      type="text"
      className="cell-edit-overlay-input"
      value={value}
      onChange={handleChange}
      onKeyDown={handleKeyDown}
      onKeyUp={handleSelect}
      onClick={handleSelect}
      onSelect={handleSelect}
      onBlur={handleBlur}
      style={{
        position: 'absolute',
        left: `${bounds.x}px`,
        top: `${bounds.y}px`,
        width: `${bounds.width}px`,
        height: `${bounds.height}px`,
        border: 'none',
        backgroundColor: '#ffffff',
        fontFamily: 'Segoe UI, Arial, sans-serif',
        fontSize: `${fontSize}px`,
        padding: '2px 4px',
        margin: 0,
        outline: 'none',
        boxSizing: 'border-box',
        zIndex: 1001,
        boxShadow: 'none',
        pointerEvents: 'auto',
        lineHeight: `${Math.max(bounds.height - 4, fontSize)}px`,
        verticalAlign: 'middle',
      }}
      spellCheck={false}
      autoComplete="off"
      data-native-text-undo
    />
  );
};
