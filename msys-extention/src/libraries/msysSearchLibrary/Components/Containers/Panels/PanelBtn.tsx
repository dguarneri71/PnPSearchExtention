import * as React from 'react';
import styles from './PanelBtn.module.scss';
import { Panel } from 'office-ui-fabric-react/lib/Panel';
import { IconButton, IIconProps } from 'office-ui-fabric-react';
import { useBoolean } from '@uifabric/react-hooks';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IPanelBtnProps {
  listID?: string;
  item?: any;
  cssUrl?: string;
}

const infoIcon: IIconProps = { iconName: 'Info' };

export const PanelBtn: React.FunctionComponent<IPanelBtnProps> = (props) => {
  const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] = useBoolean(false);

  if (props.cssUrl) {
    SPComponentLoader.loadCss(props.cssUrl);
  }

  return (
    <div>
      <IconButton title="View info" iconProps={infoIcon} onClick={openPanel}></IconButton>
      <Panel
        headerText={props.item.Filename}
        isOpen={isOpen}
        onDismiss={dismissPanel}
        // You MUST provide this prop! Otherwise screen readers will just say "button" with no label.
        closeButtonAriaLabel="Close"
      >
        <p>
          <div className={`${styles.csInfoMCLine} products`}>&nbsp;</div>
          <div className={styles.csInfoContent}>
            <div>
              <label className={styles.csInfoLabel}>Title</label>
              <div className={styles.csInfoText}>{props.item.Title}</div>
              <div>{props.children}</div>
            </div>
          </div>
        </p>
      </Panel>
    </div>
  );
};