// src/components/backstage/panels/OpenPanel.tsx

import React, { useState, useCallback, useMemo, useRef, useEffect } from 'react';
import type { FileOperations, RecentFile } from '@cyber-sheet/core';
import { loadXlsxFromArrayBuffer } from '@cyber-sheet/io-xlsx';

export interface OpenPanelProps {
  fileOperations: FileOperations;
  onWorkbookLoaded?: (workbook: any) => void;
  onOpenFile: (fileId: string) => void;
}

type FileSource = 'all' | 'onedrive' | 'local' | 'sharepoint' | 'shared';
type SortField = 'name' | 'lastModified' | 'location';
type SortDirection = 'asc' | 'desc';

const SOURCE_TABS: { id: FileSource; label: string; icon: string }[] = [
  { id: 'all', label: 'All', icon: '📂' },
  { id: 'onedrive', label: 'OneDrive', icon: '☁️' },
  { id: 'local', label: 'This PC', icon: '💻' },
  { id: 'sharepoint', label: 'SharePoint', icon: '🏢' },
  { id: 'shared', label: 'Shared with Me', icon: '👥' },
];

const LOCATION_LABELS: Record<string, string> = {
  onedrive: 'OneDrive',
  local: 'This PC',
  sharepoint: 'SharePoint',
};

const LOCATION_ICONS: Record<string, string> = {
  onedrive: '☁️',
  local: '💻',
  sharepoint: '🏢',
};

export const OpenPanel: React.FC<OpenPanelProps> = ({
  fileOperations,
  onWorkbookLoaded,
  onOpenFile,
}) => {
  const [activeSource, setActiveSource] = useState<FileSource>('all');
  const [searchQuery, setSearchQuery] = useState('');
  const [sortField, setSortField] = useState<SortField>('lastModified');
  const [sortDirection, setSortDirection] = useState<SortDirection>('desc');
  const [hoveredFile, setHoveredFile] = useState<string | null>(null);
  const [pinningFile, setPinningFile] = useState<string | null>(null);
  
  const searchInputRef = useRef(null as HTMLInputElement | null);
  const searchDebounceRef = useRef(null as NodeJS.Timeout | null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Focus search on mount
  useEffect(() => {
    searchInputRef.current?.focus();
  }, []);

  // Get all files
  const allFiles = useMemo(() => {
    const recent = fileOperations.getRecentFiles();
    const pinned = fileOperations.getPinnedFiles();
    const shared = fileOperations.getSharedFiles();
    
    // Merge, deduplicate, and mark pinned status
    const fileMap = new Map<string, RecentFile>();
    
    recent.forEach(f => fileMap.set(f.id, f));
    pinned.forEach(f => fileMap.set(f.id, { ...f, isPinned: true }));
    shared.forEach(f => fileMap.set(f.id, { ...f, isPinned: f.isPinned }));
    
    return Array.from(fileMap.values());
  }, [fileOperations]);

  // Filter by source and search
  const filteredFiles = useMemo(() => {
    let files = [...allFiles];
    
    // Source filter
    if (activeSource !== 'all') {
      if (activeSource === 'shared') {
        files = files.filter(f => f.sharedBy !== undefined);
      } else {
        files = files.filter(f => f.location === activeSource);
      }
    }
    
    // Search filter
    if (searchQuery.trim()) {
      const query = searchQuery.toLowerCase();
      files = files.filter(f => 
        f.name.toLowerCase().includes(query) ||
        f.path.toLowerCase().includes(query)
      );
    }
    
    // Sort
    files.sort((a, b) => {
      let comparison = 0;
      
      if (sortField === 'name') {
        comparison = a.name.localeCompare(b.name);
      } else if (sortField === 'lastModified') {
        comparison = a.lastModified.getTime() - b.lastModified.getTime();
      } else if (sortField === 'location') {
        comparison = a.location.localeCompare(b.location);
      }
      
      return sortDirection === 'asc' ? comparison : -comparison;
    });
    
    // Pinned files first
    const pinned = files.filter(f => f.isPinned);
    const unpinned = files.filter(f => !f.isPinned);
    
    return [...pinned, ...unpinned];
  }, [allFiles, activeSource, searchQuery, sortField, sortDirection]);

  // Group pinned and recent
  const { pinnedFiles, recentFiles } = useMemo(() => {
    return {
      pinnedFiles: filteredFiles.filter((f: RecentFile) => f.isPinned),
      recentFiles: filteredFiles.filter((f: RecentFile) => !f.isPinned),
    };
  }, [filteredFiles]);

  const handleSearch = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const value = e.target.value;
    
    if (searchDebounceRef.current) {
      clearTimeout(searchDebounceRef.current);
    }
    
    searchDebounceRef.current = setTimeout(() => {
      setSearchQuery(value);
    }, 200);
  }, []);

  const handleSortToggle = useCallback((field: SortField) => {
    if (sortField === field) {
      // Toggle sort direction
      const newDirection: SortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
      setSortDirection(newDirection);
    } else {
      setSortField(field);
      setSortDirection('desc');
    }
  }, [sortField, sortDirection]);

  const handlePinToggle = useCallback((e: React.MouseEvent, file: RecentFile) => {
    e.stopPropagation();
    setPinningFile(file.id);
    
    // Brief animation delay
    setTimeout(() => {
      if (file.isPinned) {
        fileOperations.unpinFile(file.id);
      } else {
        fileOperations.pinFile(file.id);
      }
      setPinningFile(null);
    }, 300);
  }, [fileOperations]);

  const handleFileInputChange = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || !onWorkbookLoaded) return;

    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = loadXlsxFromArrayBuffer(new Uint8Array(arrayBuffer));
      onWorkbookLoaded(workbook);
    } catch (err) {
      console.error('Failed to load file:', err);
      alert('Failed to load file. Please ensure it is a valid XLSX file.');
    }
  }, [onWorkbookLoaded]);

  const handleBrowseClick = useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  const formatDate = (date: Date): string => {
    const now = new Date();
    const diff = now.getTime() - date.getTime();
    const hours = diff / (1000 * 60 * 60);
    
    if (hours < 1) return 'Less than an hour ago';
    if (hours < 2) return '1 hour ago';
    if (hours < 24) return `${Math.floor(hours)} hours ago`;
    if (hours < 48) return 'Yesterday';
    if (hours < 168) return `${Math.floor(hours / 24)} days ago`;
    if (hours < 720) return `${Math.floor(hours / 168)} weeks ago`;
    
    return date.toLocaleDateString('en-US', { 
      month: 'short', 
      day: 'numeric',
      year: date.getFullYear() !== now.getFullYear() ? 'numeric' : undefined,
    });
  };

  const formatFileSize = (bytes?: number): string => {
    if (!bytes) return '';
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  // Cleanup
  useEffect(() => {
    return () => {
      if (searchDebounceRef.current) {
        clearTimeout(searchDebounceRef.current);
      }
    };
  }, []);

  // ─── Styles ──────────────────────────────────────────────

  const containerStyle: React.CSSProperties = {
    padding: '32px 48px',
    maxWidth: 720,
    animation: 'fadeIn 200ms ease-out',
  };

  const headingStyle: React.CSSProperties = {
    fontSize: 24,
    fontWeight: 600,
    color: '#1F1F1F',
    marginBottom: 24,
  };

  // Search bar
  const searchContainerStyle: React.CSSProperties = {
    position: 'relative',
    marginBottom: 20,
  };

  const searchIconStyle: React.CSSProperties = {
    position: 'absolute',
    left: 12,
    top: '50%',
    transform: 'translateY(-50%)',
    fontSize: 16,
    color: '#999999',
    pointerEvents: 'none',
  };

  const searchInputStyle: React.CSSProperties = {
    width: '100%',
    padding: '10px 12px 10px 36px',
    fontSize: 14,
    border: '1px solid #D1D1D1',
    borderRadius: 20,
    outline: 'none',
    transition: 'border-color 150ms, box-shadow 150ms',
    boxSizing: 'border-box',
  };

  // Source tabs
  const tabsContainerStyle: React.CSSProperties = {
    display: 'flex',
    gap: 4,
    marginBottom: 20,
    borderBottom: '1px solid #E8E8E8',
    paddingBottom: 0,
  };

  const tabStyle = (isActive: boolean): React.CSSProperties => ({
    padding: '8px 16px',
    fontSize: 13,
    fontWeight: isActive ? 600 : 400,
    color: isActive ? '#0078D4' : '#666666',
    backgroundColor: 'transparent',
    border: 'none',
    borderBottom: isActive ? '2px solid #0078D4' : '2px solid transparent',
    cursor: 'pointer',
    transition: 'all 150ms',
    display: 'flex',
    alignItems: 'center',
    gap: 6,
    marginBottom: -1,
  });

  // Sort header
  const sortHeaderStyle: React.CSSProperties = {
    display: 'grid',
    gridTemplateColumns: '1fr 160px 100px 40px',
    gap: 12,
    padding: '8px 12px',
    borderBottom: '1px solid #E8E8E8',
    marginBottom: 4,
  };

  const sortHeaderItemStyle = (field: SortField): React.CSSProperties => ({
    fontSize: 11,
    fontWeight: 600,
    color: sortField === field ? '#0078D4' : '#888888',
    textTransform: 'uppercase' as const,
    letterSpacing: '0.5px',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    gap: 4,
    userSelect: 'none' as const,
  });

  // File rows
  const sectionLabelStyle: React.CSSProperties = {
    fontSize: 12,
    fontWeight: 600,
    color: '#888888',
    textTransform: 'uppercase' as const,
    letterSpacing: '0.5px',
    padding: '16px 12px 8px',
  };

  const emptyStateStyle: React.CSSProperties = {
    textAlign: 'center' as const,
    padding: '60px 20px',
    color: '#999999',
  };

  const fileListContainerStyle: React.CSSProperties = {
    maxHeight: 'calc(100vh - 320px)',
    overflowY: 'auto' as const,
  };

  const browseButtonStyle: React.CSSProperties = {
    display: 'inline-flex',
    alignItems: 'center',
    gap: 8,
    padding: '8px 16px',
    fontSize: 14,
    fontWeight: 500,
    color: '#FFFFFF',
    backgroundColor: '#0078D4',
    border: 'none',
    borderRadius: 4,
    cursor: 'pointer',
    transition: 'background-color 150ms',
    marginBottom: 16,
  };

  return (
    <div style={containerStyle}>
      <h1 style={headingStyle}>Open</h1>

      {/* Browse from computer */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls"
        onChange={handleFileInputChange}
        style={{ display: 'none' }}
      />
      <button
        onClick={handleBrowseClick}
        style={browseButtonStyle}
        onMouseEnter={(e: React.MouseEvent<HTMLButtonElement>) => {
          e.currentTarget.style.backgroundColor = '#005A9E';
        }}
        onMouseLeave={(e: React.MouseEvent<HTMLButtonElement>) => {
          e.currentTarget.style.backgroundColor = '#0078D4';
        }}
      >
        📁 Browse from Computer
      </button>

      {/* Search */}
      <div style={searchContainerStyle}>
        <span style={searchIconStyle}>🔍</span>
        <input
          ref={searchInputRef}
          type="text"
          placeholder="Search files..."
          onChange={handleSearch}
          style={searchInputStyle}
          onFocus={(e: React.FocusEvent<HTMLInputElement>) => {
            e.currentTarget.style.borderColor = '#0078D4';
            e.currentTarget.style.boxShadow = '0 0 0 3px rgba(0,120,212,0.15)';
          }}
          onBlur={(e: React.FocusEvent<HTMLInputElement>) => {
            e.currentTarget.style.borderColor = '#D1D1D1';
            e.currentTarget.style.boxShadow = 'none';
          }}
          aria-label="Search files"
        />
      </div>

      {/* Source tabs */}
      <div style={tabsContainerStyle} role="tablist">
        {SOURCE_TABS.map(tab => (
          <button
            key={tab.id}
            style={tabStyle(activeSource === tab.id)}
            onClick={() => setActiveSource(tab.id)}
            role="tab"
            aria-selected={activeSource === tab.id}
            onMouseEnter={(e: React.MouseEvent<HTMLButtonElement>) => {
              if (activeSource !== tab.id) {
                e.currentTarget.style.color = '#333333';
              }
            }}
            onMouseLeave={(e: React.MouseEvent<HTMLButtonElement>) => {
              if (activeSource !== tab.id) {
                e.currentTarget.style.color = '#666666';
              }
            }}
          >
            <span>{tab.icon}</span>
            {tab.label}
          </button>
        ))}
      </div>

      {/* Sort header */}
      <div style={sortHeaderStyle}>
        <div
          style={sortHeaderItemStyle('name')}
          onClick={() => handleSortToggle('name')}
        >
          Name
          {sortField === 'name' && (
            <span>{sortDirection === 'asc' ? '↑' : '↓'}</span>
          )}
        </div>
        <div
          style={sortHeaderItemStyle('lastModified')}
          onClick={() => handleSortToggle('lastModified')}
        >
          Date Modified
          {sortField === 'lastModified' && (
            <span>{sortDirection === 'asc' ? '↑' : '↓'}</span>
          )}
        </div>
        <div
          style={sortHeaderItemStyle('location')}
          onClick={() => handleSortToggle('location')}
        >
          Location
          {sortField === 'location' && (
            <span>{sortDirection === 'asc' ? '↑' : '↓'}</span>
          )}
        </div>
        <div style={{ width: 40 }} />
      </div>

      {/* File list */}
      <div style={fileListContainerStyle}>
        {filteredFiles.length === 0 ? (
          <div style={emptyStateStyle}>
            <div style={{ fontSize: 40, marginBottom: 12 }}>📭</div>
            <div style={{ fontSize: 15, fontWeight: 500, marginBottom: 4 }}>
              No files found
            </div>
            <div style={{ fontSize: 13 }}>
              {searchQuery 
                ? `No files matching "${searchQuery}"`
                : 'No recent files in this location'}
            </div>
          </div>
        ) : (
          <>
            {/* Pinned files section */}
            {pinnedFiles.length > 0 && (
              <>
                <div style={sectionLabelStyle}>📌 Pinned</div>
                {pinnedFiles.map((file: RecentFile) => (
                  <div key={file.id}>
                    <FileRow
                      file={file}
                      isHovered={hoveredFile === file.id}
                      isPinning={pinningFile === file.id}
                      onMouseEnter={() => setHoveredFile(file.id)}
                      onMouseLeave={() => setHoveredFile(null)}
                      onClick={() => onOpenFile(file.id)}
                      onPinToggle={(e) => handlePinToggle(e, file)}
                      formatDate={formatDate}
                      formatFileSize={formatFileSize}
                    />
                  </div>
                ))}
              </>
            )}

            {/* Recent files section */}
            {recentFiles.length > 0 && (
              <>
                {pinnedFiles.length > 0 && (
                  <div style={sectionLabelStyle}>Recent</div>
                )}
                {recentFiles.map((file: RecentFile) => (
                  <div key={file.id}>
                    <FileRow
                      file={file}
                      isHovered={hoveredFile === file.id}
                      isPinning={pinningFile === file.id}
                      onMouseEnter={() => setHoveredFile(file.id)}
                      onMouseLeave={() => setHoveredFile(null)}
                      onClick={() => onOpenFile(file.id)}
                      onPinToggle={(e) => handlePinToggle(e, file)}
                      formatDate={formatDate}
                      formatFileSize={formatFileSize}
                    />
                  </div>
                ))}
              </>
            )}
          </>
        )}
      </div>

      <style>{`
        @keyframes fadeIn {
          from { opacity: 0; }
          to { opacity: 1; }
        }
      `}</style>
    </div>
  );
};

// ─── FileRow Sub-component ─────────────────────────────────

interface FileRowProps {
  file: RecentFile;
  isHovered: boolean;
  isPinning: boolean;
  onMouseEnter: () => void;
  onMouseLeave: () => void;
  onClick: () => void;
  onPinToggle: (e: React.MouseEvent) => void;
  formatDate: (date: Date) => string;
  formatFileSize: (bytes?: number) => string;
}

const FileRow: React.FC<FileRowProps> = ({
  file,
  isHovered,
  isPinning,
  onMouseEnter,
  onMouseLeave,
  onClick,
  onPinToggle,
  formatDate,
  formatFileSize,
}) => {
  const locationIcon = LOCATION_ICONS[file.location] || '📁';
  const locationLabel = LOCATION_LABELS[file.location] || file.location;

  return (
    <div
      style={{
        display: 'grid',
        gridTemplateColumns: '1fr 160px 100px 40px',
        gap: 12,
        padding: '10px 12px',
        borderRadius: 6,
        cursor: 'pointer',
        backgroundColor: isHovered ? '#F5F5F5' : 'transparent',
        transition: 'background-color 100ms',
        alignItems: 'center',
      }}
      onClick={onClick}
      onMouseEnter={onMouseEnter}
      onMouseLeave={onMouseLeave}
      role="button"
      tabIndex={0}
      onKeyDown={(e: React.KeyboardEvent<HTMLDivElement>) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          onClick();
        }
      }}
      aria-label={`Open ${file.name}`}
    >
      {/* File name */}
      <div style={{
        fontSize: 14,
        fontWeight: 500,
        color: '#1F1F1F',
        display: 'flex',
        alignItems: 'center',
        gap: 8,
        overflow: 'hidden',
        textOverflow: 'ellipsis',
        whiteSpace: 'nowrap',
      }}>
        <span style={{ fontSize: 20, flexShrink: 0 }}>📄</span>
        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>
          {file.name}
        </span>
      </div>

      {/* Date modified */}
      <div style={{ fontSize: 12, color: '#888888' }}>
        {formatDate(file.lastModified)}
      </div>

      {/* Location */}
      <div style={{ 
        fontSize: 12, 
        color: '#888888',
        display: 'flex',
        alignItems: 'center',
        gap: 4,
      }}>
        <span>{locationIcon}</span>
        <span>{locationLabel}</span>
      </div>

      {/* Pin button */}
      <button
        style={{
          width: 32,
          height: 32,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          border: 'none',
          backgroundColor: isHovered ? '#E8E8E8' : 'transparent',
          cursor: 'pointer',
          borderRadius: 4,
          fontSize: 18,
          color: file.isPinned ? '#0078D4' : '#CCCCCC',
          transform: isPinning ? 'scale(1.3) rotate(45deg)' : file.isPinned ? 'rotate(0deg)' : 'rotate(-45deg)',
          transition: 'all 250ms ease',
          opacity: isHovered || file.isPinned ? 1 : 0,
        }}
        onClick={onPinToggle}
        onMouseEnter={(e: React.MouseEvent<HTMLButtonElement>) => {
          e.currentTarget.style.backgroundColor = '#D0D0D0';
        }}
        onMouseLeave={(e: React.MouseEvent<HTMLButtonElement>) => {
          e.currentTarget.style.backgroundColor = isHovered ? '#E8E8E8' : 'transparent';
        }}
        title={file.isPinned ? 'Unpin from top' : 'Pin to top'}
        aria-label={file.isPinned ? 'Unpin file' : 'Pin file'}
      >
        📌
      </button>
    </div>
  );
};
