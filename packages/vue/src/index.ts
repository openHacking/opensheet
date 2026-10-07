import { defineComponent, h, onMounted, onBeforeUnmount, watch, ref, type PropType } from 'vue';
import { createOpenSheet, type OpenSheet, type Commit, type WorkbookFile } from 'opensheet';
export const OpenSheetView = defineComponent({
  name: 'OpenSheetView',
  props: { initialFile: Object as PropType<WorkbookFile>, readOnly: Boolean },
  emits: ['ready', 'change', 'error'],
  setup(props, { emit, expose }) {
    const host = ref<HTMLElement>();
    let app: OpenSheet | undefined;
    let active = true;
    onMounted(async () => {
      try {
        app = createOpenSheet({
          container: host.value!,
          mode: props.readOnly ? 'read' : 'edit',
          onError: (error) => emit('error', error),
        });
        if (props.initialFile) await app.load(props.initialFile);
        else await app.createWorkbook();
        if (!active) return;
        app.on('workbook:committed', (commit) => emit('change', commit));
        emit('ready', app);
      } catch (error) {
        if (active) emit('error', error);
      }
    });
    watch(
      () => props.readOnly,
      (v) => app?.setMode(v ? 'read' : 'edit'),
    );
    onBeforeUnmount(() => {
      active = false;
      app?.dispose();
    });
    expose({ getInstance: () => app });
    return () => h('div', { ref: host, style: { height: '500px' } });
  },
});
