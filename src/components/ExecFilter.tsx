import * as React from 'react'

import { Persona } from 'office-ui-fabric-react/lib/Persona'

import { IExecutive,cdnAssetsBaseUrl } from '../shared/SharePoint'

import styles from './ExecFilter.module.scss';

export interface IExecFilterProps {
  execs: IExecutive[]
  onChange: (execs: number[]) => void
}

export interface IExecFilterState {
  execs: IExecutive2[]
  collapsed: boolean
  testExecs: any[]
}

export interface IExecutive2 extends IExecutive {
  active: boolean
}

export default class ExecFilter extends React.Component<IExecFilterProps, IExecFilterState> {
  private initalFilterCount = 3
  
  constructor(props) {
    super(props)
    this.state = {
      execs: this.props.execs.map(exec => {
        return {
          ...exec, 
          active: true
        }
      }),
      collapsed: true,
      testExecs: this.props.execs,
    }
  }

  public componentWillReceiveProps(nextProps) {
    if (nextProps.execs.length) {
      this.setState({
        execs: nextProps.execs,
      })
    }
  }
  private filter(id){
    let activeCount = this.getActiveFilterCount()
    const len = this.state.execs.length
    for (let i = 0; i < len; i += 1) {
      const exec = this.state.execs[i]
      if (exec.id === id) {

        //TODO: call list function filter function with id(s)
        if(exec.active && (activeCount === len)) {
          //exec.active = false
          this.filterBySingleExec(exec.id)
        }
        else if(exec.active && (activeCount <= len && activeCount > 1)) {
          exec.active = false
        }
        else if(exec.active && activeCount === 1) {
          this.resetFilter()
        }
        else if(!exec.active) {
          exec.active = true
        }        
        this.props.onChange(this.getActiveExecs())
        this.setState({ execs: this.state.execs })
        break
      }
    }
  }

  private getActiveExecs() {
    return this.state.execs
      .filter(exec => exec.active)
      .map(exec => exec.id)
  }

  private getActiveFilterCount() {
    return this.state.execs
      .filter(exec => exec.active)
      .length
  }

  private filterBySingleExec(id) {
    this.state.execs.forEach(exec => {
      if (exec.id !== id) {
        exec.active = false
      } 
    })
    this.setState({ execs: this.state.execs })
  }

  private resetFilter() {
    this.state.execs.forEach(exec =>{
      exec.active = true;
      this.props.onChange(this.getActiveExecs())
      this.setState({ execs: this.state.execs })
    })
  }

  private toggleMore(){
    this.setState({ collapsed: !this.state.collapsed })
  }

  public render() {
    const { initalFilterCount } = this
    const { collapsed, execs } = this.state

    const execsFiltered = collapsed ? execs.slice(0, initalFilterCount) : execs;
    const moreButtonText = collapsed ? "More" : "Less"
    const moreButtonContent = collapsed ? styles.more : styles.less
    

    return (
      <div className={styles.ExecFilter}>
        <div className={styles.personas}>
          {execsFiltered.map((execFilter, i) => {
            return (
              <div key={i} onClick={() => this.filter(execFilter.id)}>
                <Persona 
                  hidePersonaDetails={false}
                  imageInitials={execFilter.initials}
                  primaryText={execFilter.name}
                  imageUrl={execFilter.imageUrl==""?`${cdnAssetsBaseUrl}/images/DefaultProfile.png` :execFilter.imageUrl}
                  className={!execFilter.active ? styles.inactive : ""}
                />
                
              </div>
            )
          }
          )}
          <div className={styles.buttons}>
            <div className={execs.length <= initalFilterCount ? styles.hidden : ""}
              onClick={() => this.toggleMore()}>
              <Persona 
                className={`${styles.moreButton} ${moreButtonContent}`}
                imageInitials= "&#xfeff;"
                primaryText={moreButtonText}
              />
            </div>
            <div className={styles.clearButton} onClick={() => this.resetFilter()}>
              Clear
            </div>
          </div>
        </div>
      </div>
    )
  }
}
