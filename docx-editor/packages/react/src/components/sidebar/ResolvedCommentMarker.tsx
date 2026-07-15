import type { Comment } from '@npde/docx-editor-core/types/content';
import { MaterialSymbol } from '../ui/Icons';
import { useTranslation } from '../../i18n';
import type { SidebarItemRenderProps } from '../../plugin-api/types';

export interface ResolvedCommentMarkerProps extends SidebarItemRenderProps {
  comment: Comment;
}

export function ResolvedCommentMarker({
  comment,
  measureRef,
  onToggleExpand,
}: ResolvedCommentMarkerProps) {
  const { t } = useTranslation();

  return (
    <div
      ref={measureRef}
      data-comment-id={comment.id}
      data-comment-resolved="true"
      role="button"
      tabIndex={0}
      aria-label={t('commentMarkers.resolvedComment')}
      onClick={onToggleExpand}
      onKeyDown={(e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          onToggleExpand();
        }
      }}
      onMouseDown={(e) => e.stopPropagation()}
      style={{
        display: 'inline-flex',
        alignItems: 'center',
        cursor: 'pointer',
        color: 'var(--doc-text-muted)',
        padding: 2,
      }}
      onMouseOver={(e) => {
        (e.currentTarget as HTMLElement).style.opacity = '0.7';
      }}
      onMouseOut={(e) => {
        (e.currentTarget as HTMLElement).style.opacity = '1';
      }}
    >
      <MaterialSymbol name="chat_bubble_check" size={20} />
    </div>
  );
}
