import * as React from 'react'

import {
  Persona
} from 'office-ui-fabric-react/lib/Persona'

import {
  IExecutive,
  cdnAssetsBaseUrl
} from '../shared/SharePoint'

import styles from './ExecFilter.module.scss';

export interface IActiveExecutive extends IExecutive {
  active: boolean
}

export interface IExecSelectorProps {
  execs: IExecutive[]
  selectedExecIds?: number[]
  onChange?: (execs: IExecutive[]) => void
  disabled?: boolean
}

export interface IExecSelectorState {
  execs: IActiveExecutive[]
}

export default class ExecSelector extends React.Component<IExecSelectorProps, IExecSelectorState> {
  constructor(props: IExecSelectorProps) {
    super(props)
    
    // console.log("ExecSelector props.execs:", props.execs)

    const ids = props.selectedExecIds || []
    const isActive = id => (
      !!ids.length && ids.indexOf(id) >= 0
    )

    this.state = {
      execs: this.props.execs.map(exec => {
        return {
          ...exec, 
          active: isActive(exec.id)
        }
      })
    }
  }

  public render() {
    const { execs } = this.state
    
    return (
      <div className={styles.ExecFilter}>
        <div className={styles.personas}>
          {execs.map((exec, i) => (
            <div key={i} onClick={() => this.toggle(i)}>
              <Persona 
                hidePersonaDetails={false}
                imageInitials={exec.initials}
                primaryText={exec.name}
                imageUrl={exec.imageUrl==""?`${cdnAssetsBaseUrl}/images/DefaultProfile.png` :exec.imageUrl}
                className={!exec.active ? styles.inactive : ""}
              />
            </div>
          ))}
        </div>
      </div>
    )
  }

  private toggle(index: number) {
    if (!this.props.disabled) {
      const { execs } = this.state
      execs[index].active = !execs[index].active
      this.onChange()
    }
  }

  // private clear() {
  //   const { execs } = this.state

  //   execs.forEach(exec => exec.active = false)
  //   this.onChange()
  // }

  private onChange() {
    const { onChange } = this.props
    
    if (onChange) {
      const { execs } = this.state
      const selectedExecs = execs.filter(exec => exec.active)
      
      onChange(selectedExecs as IExecutive[])
      this.forceUpdate()
    }
  }
}
