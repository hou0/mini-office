import { useCallback, useEffect, useRef, useState } from 'react';
import { useWorkspaceStore, type WordSession, type SheetSession, type SlidesSession } from '@mini-office/document-model';
import { applyAppearance, watchSystemTheme } from '@mini-office/design-system';
import { ToastViewport, TooltipHost } from '@mini-office/ui';
import { WordEditor } from '@mini-office/editor-word';
import { SheetEditor } from '@mini-office/editor-sheet';
import { SlidesEditor } from '@mini-office/editor-slides';
import { extensionOf } from '@mini-office/shared';
import {
  BusyIndicator,
  CloseConfirmDialog,
  DropOverlay,
  EditorLoading,
  StatusBar,
  TabBar,
  TitleBar,
} from './chrome';
import { CommandPalette } from './CommandPalette';
import { WorkspaceDialogs } from './dialogs';
import { HomePage } from './HomePage';

export function AppShell() {
  const sessions = useWorkspaceStore((state) => state.sessions);
  const tabOrder = useWorkspaceStore((state) => state.tabOrder);
  const activeId = useWorkspaceStore((state) => state.activeId);
  const settings = useWorkspaceStore((state) => state.settings);
  const dialog = useWorkspaceStore((state) => state.dialog);
  const closeTab = useWorkspaceStore((state) => state.closeTab);
  const saveSession = useWorkspaceStore((state) => state.saveSession);
  const saveActive = useWorkspaceStore((state) => state.saveActive);
  const openFileDialog = useWorkspaceStore((state) => state.openFileDialog);
  const openDroppedFile = useWorkspaceStore((state) => state.openDroppedFile);
  const createDocument = useWorkspaceStore((state) => state.createDocument);
  const setCommandPalette = useWorkspaceStore((state) => state.setCommandPalette);
  const commandPaletteOpen = useWorkspaceStore((state) => state.commandPaletteOpen);
  const closeDialog = useWorkspaceStore((state) => state.closeDialog);
  const openDialog = useWorkspaceStore((state) => state.openDialog);
  const notifications = useWorkspaceStore((state) => state.notifications);
  const dismissNotification = useWorkspaceStore((state) => state.dismissNotification);
  const notify = useWorkspaceStore((state) => state.notify);

  const [showHome, setShowHome] = useState(!activeId);
  const [closeRequestId, setCloseRequestId] = useState<string | null>(null);
  const [closing, setClosing] = useState(false);
  const [dragging, setDragging] = useState(false);
  const previousActive = useRef<string | null>(activeId);
  const dragDepth = useRef(0);

  const activeSession = activeId ? sessions[activeId] : undefined;
  const closeRequestSession = closeRequestId ? sessions[closeRequestId] : undefined;

  const requestClose = useCallback(
    (id: string) => {
      const session = useWorkspaceStore.getState().sessions[id];
      if (!session) return;
      if (session.dirty) setCloseRequestId(id);
      else void useWorkspaceStore.getState().closeTab(id);
    },
    [],
  );

  useEffect(() => {
    if (activeId && activeId !== previousActive.current) setShowHome(false);
    if (!activeId) setShowHome(true);
    previousActive.current = activeId;
  }, [activeId]);

  useEffect(() => {
    applyAppearance(settings);
    return watchSystemTheme(settings.theme);
  }, [settings.theme, settings.accent, settings.customAccent]);

  useEffect(() => {
    document.documentElement.dataset.fontSize = settings.fontSize;
  }, [settings.fontSize]);

  useEffect(() => {
    document.title = activeSession ? `${activeSession.displayName} — Mini Office` : 'Mini Office';
  }, [activeSession]);

  useEffect(() => {
    const hasDirty = Object.values(sessions).some((session) => session.dirty);
    const handler = (event: BeforeUnloadEvent) => {
      if (!hasDirty) return;
      event.preventDefault();
      event.returnValue = '';
    };
    window.addEventListener('beforeunload', handler);
    return () => window.removeEventListener('beforeunload', handler);
  }, [sessions]);

  useEffect(() => {
    const handler = (event: KeyboardEvent) => {
      const mod = event.metaKey || event.ctrlKey;
      const target = event.target as HTMLElement | null;
      const inEditable =
        !!target &&
        (target.tagName === 'INPUT' ||
          target.tagName === 'TEXTAREA' ||
          target.tagName === 'SELECT' ||
          target.isContentEditable);
      if (mod && event.key.toLowerCase() === 'k') {
        event.preventDefault();
        setCommandPalette(!commandPaletteOpen);
        return;
      }
      if (mod && event.key.toLowerCase() === 's') {
        event.preventDefault();
        if (activeId) {
          if (event.shiftKey) void saveActive({ asNew: true });
          else void saveActive();
        }
        return;
      }
      if (mod && event.key.toLowerCase() === 'o') {
        event.preventDefault();
        void openFileDialog();
        return;
      }
      if (mod && event.key.toLowerCase() === 'n' && !event.shiftKey) {
        event.preventDefault();
        void createDocument('docx');
        return;
      }
      if (mod && event.key === ',') {
        event.preventDefault();
        openDialog({ type: 'settings' });
        return;
      }
      if (mod && event.key.toLowerCase() === 'w') {
        event.preventDefault();
        if (activeId) requestClose(activeId);
        return;
      }
      if (mod && event.key.toLowerCase() === 'p') {
        event.preventDefault();
        if (activeId) window.dispatchEvent(new CustomEvent('mo:print'));
        return;
      }
      if (mod && event.key.toLowerCase() === 'f') {
        event.preventDefault();
        if (activeId) window.dispatchEvent(new CustomEvent('mo:find'));
        return;
      }
      if (event.key === 'F5') {
        event.preventDefault();
        if (activeSession?.kind === 'pptx') window.dispatchEvent(new CustomEvent('mo:present'));
        return;
      }
      if (event.key === 'Escape') {
        if (commandPaletteOpen) {
          setCommandPalette(false);
          return;
        }
        if (dialog) {
          closeDialog();
          return;
        }
        if (closeRequestId) {
          setCloseRequestId(null);
          return;
        }
      }
      if (mod && !inEditable && /^[1-9]$/.test(event.key)) {
        const index = Number(event.key) - 1;
        const id = tabOrder[index];
        if (id) {
          event.preventDefault();
          useWorkspaceStore.getState().activateTab(id);
          setShowHome(false);
        }
      }
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [
    activeId,
    activeSession?.kind,
    closeRequestId,
    closeDialog,
    commandPaletteOpen,
    createDocument,
    dialog,
    openDialog,
    openFileDialog,
    saveActive,
    setCommandPalette,
    tabOrder,
    requestClose,
  ]);

  const discardAndClose = async () => {
    if (!closeRequestId) return;
    await closeTab(closeRequestId, true);
    setCloseRequestId(null);
  };

  const saveAndClose = async () => {
    if (!closeRequestId) return;
    setClosing(true);
    const saved = await saveSession(closeRequestId);
    setClosing(false);
    if (saved) {
      await closeTab(closeRequestId, true);
      setCloseRequestId(null);
    }
  };

  const handleDrop = async (event: React.DragEvent) => {
    event.preventDefault();
    dragDepth.current = 0;
    setDragging(false);
    const files = Array.from(event.dataTransfer.files);
    if (files.length === 0) return;
    const supported = files.find((file) =>
      ['.docx', '.docm', '.xlsx', '.xlsm', '.pptx', '.pptm', '.csv', '.doc', '.xls', '.ppt'].includes(
        extensionOf(file.name),
      ),
    );
    if (!supported) {
      notify('warning', '不支持的文件类型', '请选择 DOCX、XLSX、PPTX 或 CSV 文件。');
      return;
    }
    if (files.length > 1) notify('info', '一次打开一个文件', `已选择 ${supported.name}`);
    await openDroppedFile(supported);
  };

  const content = (() => {
    if (showHome || !activeSession) return <HomePage />;
    if (activeSession.state === 'VALIDATING' || activeSession.state === 'LOADING') {
      return <EditorLoading title={activeSession.displayName} />;
    }
    if (activeSession.kind === 'docx') {
      return <WordEditor key={activeSession.id} session={activeSession as WordSession} />;
    }
    if (activeSession.kind === 'xlsx') {
      return <SheetEditor key={activeSession.id} session={activeSession as SheetSession} />;
    }
    return <SlidesEditor key={activeSession.id} session={activeSession as SlidesSession} />;
  })();

  return (
    <div
      className="mo-app flex h-full flex-col overflow-hidden bg-[var(--mo-bg)]"
      onDragEnter={(event) => {
        if (!event.dataTransfer.types.includes('Files')) return;
        dragDepth.current += 1;
        setDragging(true);
      }}
      onDragOver={(event) => {
        if (event.dataTransfer.types.includes('Files')) event.preventDefault();
      }}
      onDragLeave={() => {
        dragDepth.current = Math.max(0, dragDepth.current - 1);
        if (dragDepth.current === 0) setDragging(false);
      }}
      onDrop={(event) => void handleDrop(event)}
    >
      <TitleBar
        onHome={() => setShowHome(true)}
        onOpenPalette={() => setCommandPalette(true)}
      />
      {tabOrder.length > 0 ? (
        <TabBar
          onRequestClose={requestClose}
          onHome={() => setShowHome(true)}
          onSelectTab={() => setShowHome(false)}
        />
      ) : null}
      <main className="min-h-0 flex-1">{content}</main>
      <StatusBar />
      <BusyIndicator />
      <DropOverlay visible={dragging} />
      <CommandPalette
        open={commandPaletteOpen}
        onClose={() => setCommandPalette(false)}
        onHome={() => {
          setShowHome(true);
          useWorkspaceStore.getState().closeDialog();
        }}
        onRequestClose={requestClose}
      />
      <WorkspaceDialogs />
      {closeRequestSession ? (
        <CloseConfirmDialog
          session={closeRequestSession}
          saving={closing}
          onCancel={() => setCloseRequestId(null)}
          onDiscard={() => void discardAndClose()}
          onSave={() => void saveAndClose()}
        />
      ) : null}
      <TooltipHost />
      <ToastViewport
        toasts={notifications.map((notification) => ({
          id: notification.id,
          level: notification.level,
          title: notification.title,
          message: notification.message,
        }))}
        onDismiss={dismissNotification}
      />
    </div>
  );
}
