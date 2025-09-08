import {Disposable, Observable, dom, input, MultiHolder, Computed} from 'grainjs';
import {saveModal} from 'app/client/ui2018/modals';
import {SheetCreationUtils} from 'app/client/lib/SheetCreationUtils';

export function showNewSheetNameModal(onSubmit: (name: string) => Promise<void> | void, defaultName = '', gristDoc?: any) {
  saveModal((_ctl, owner): any => {
    // Generate the auto-generated unique sheet name if gristDoc is provided
    let autoName = defaultName;
    if (gristDoc && defaultName) {
      try {
        autoName = SheetCreationUtils._generateUniqueTableName(gristDoc, defaultName);
      } catch (e) {
        autoName = defaultName;
      }
    }
    const sheetName = Observable.create(owner, autoName);
    return {
      title: 'Name your new sheet',
      body: dom('div',
        dom('label', 'Sheet Name:'),
        input(sheetName, dom.attr('placeholder', 'Enter sheet name'), dom.style('width', '100%'))
      ),
      saveLabel: 'Create',
      saveFunc: async () => {
        const name = sheetName.get().trim();
        if (!name) { throw new Error('Sheet name cannot be empty'); }
        await onSubmit(name);
      },
      saveDisabled: Computed.create(owner, (use) => !use(sheetName).trim()),
    };
  });
}
