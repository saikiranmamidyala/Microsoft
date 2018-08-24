import * as React from 'react'

import {
  Spinner,
  SpinnerSize
} from 'office-ui-fabric-react/lib/Spinner'

import styles from './Loading.module.scss'

export default class Loading extends React.PureComponent {
  public render(): JSX.Element {
    return (
      <div className={styles.Loading}>
        <Spinner size={SpinnerSize.large} />
      </div>
    )
  }
}
