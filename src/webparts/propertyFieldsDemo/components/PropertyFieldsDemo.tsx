import * as React from 'react';
import styles from './PropertyFieldsDemo.module.scss';
import { IPropertyFieldsDemoProps } from './IPropertyFieldsDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class PropertyFieldsDemo extends React.Component<IPropertyFieldsDemoProps, {}> {
    public render(): React.ReactElement<IPropertyFieldsDemoProps> {
        return (
            <div className={styles.propertyFieldsDemo}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <span className={styles.title}>Welcome to SharePoint!</span>
                            <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
                            <p className={styles.description}>Selected List: {this.props.list}</p>
                            <p className={styles.description}>Single Column ID: {this.props.column}</p>
                            <p className={styles.description}>Single Column Title(s): {this.props.columnTitle}</p>
                            <p className={styles.description}>Single Column Internal Name(s): {this.props.columnInternalName}</p>
                            <a href="https://aka.ms/spfx" className={styles.button}>
                                <span className={styles.label}>Learn more</span>
                            </a>
                        </div>
                    </div>
                </div>
            </div>
        );
    }
}
