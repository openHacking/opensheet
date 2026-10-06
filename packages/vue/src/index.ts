import { defineComponent, h, onMounted, onBeforeUnmount, watch, ref, type PropType } from 'vue';
import { createOpenSheet, type OpenSheet, type WorkbookSnapshot } from 'opensheet';
export const OpenSheetView = defineComponent({
  name: 'OpenSheetView',
  props: { initialSnapshot: Object as PropType<WorkbookSnapshot>, readOnly: Boolean },
  emits: ['ready', 'change'],
  setup(props, { emit, expose }) {
    const host = ref<HTMLElement>();
    let app: OpenSheet | undefined;
    let active = true;
    onMounted(async () => {
      app = createOpenSheet({ container: host.value!, mode: props.readOnly ? 'read' : 'edit' });
      if (props.initialSnapshot) await app.load(props.initialSnapshot);
      else app.createWorkbook();
      if (!active) return;
      app.on('workbook:committed', () => emit('change', app!.getWorkbook().toJSON()));
      emit('ready', app);
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
