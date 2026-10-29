/**
 * 功能：保存「我的样式集」对话框，把本文档的自定义样式沉淀为跨文档复用的样式包；时间：2026-10-29；作者：lq。
 */
import { useEffect, useState } from 'react';
import { Button, Dialog, Field, Input } from '@mini-office/ui';

export function SaveStyleSetDialog({
  open,
  defaultName,
  styleCount,
  existingNames,
  onClose,
  onSave,
}: {
  open: boolean;
  defaultName: string;
  styleCount: number;
  existingNames: string[];
  onClose: () => void;
  onSave: (name: string) => void;
}) {
  const [name, setName] = useState(defaultName);

  useEffect(() => {
    if (open) setName(defaultName);
  }, [defaultName, open]);

  const trimmed = name.trim();
  const overwrite = existingNames.includes(trimmed);
  const invalid = !trimmed || trimmed.length > 40;

  const submit = () => {
    if (invalid) return;
    onSave(trimmed);
    onClose();
  };

  return (
    <Dialog
      open={open}
      onClose={onClose}
      title="保存为我的样式集"
      description={`把本文档的 ${styleCount} 个自定义样式保存到本机，之后可在任意文档的样式库里一键套用。`}
      width="sm"
      footer={
        <>
          <Button onClick={onClose}>取消</Button>
          <Button variant="primary" disabled={invalid} onClick={submit}>
            {overwrite ? '覆盖保存' : '保存'}
          </Button>
        </>
      }
    >
      <Field
        label="样式集名称"
        hint={overwrite ? '已存在同名样式集，保存后将覆盖原样式集。' : '建议使用文档类型命名，例如「公文」「合同」'}
      >
        <Input
          autoFocus
          value={name}
          maxLength={60}
          onChange={(event) => setName(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              submit();
            }
          }}
          placeholder="例如：公文样式集"
        />
      </Field>
    </Dialog>
  );
}
